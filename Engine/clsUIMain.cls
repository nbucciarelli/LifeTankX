VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================
'User Interface for tab : Main
'===========================================
Option Explicit

Private InterfaceName As String

Public stVersion As DecalControls.StaticText
Public lblStatus As DecalControls.StaticText
Public lblBuyTarget As DecalControls.StaticText
Public lblSellTarget As DecalControls.StaticText

Public WithEvents chkEnable As DecalControls.CheckBox
Attribute chkEnable.VB_VarHelpID = -1

Public WithEvents btnPause As PushButton
Attribute btnPause.VB_VarHelpID = -1
Public WithEvents chCombatType As DecalControls.Choice
Attribute chCombatType.VB_VarHelpID = -1

Public WithEvents chMacroProfile As DecalControls.Choice
Attribute chMacroProfile.VB_VarHelpID = -1
Public WithEvents chBuffsProfile As DecalControls.Choice
Attribute chBuffsProfile.VB_VarHelpID = -1
Public WithEvents chLootProfile As DecalControls.Choice
Attribute chLootProfile.VB_VarHelpID = -1

Public txtMacroProfile As DecalControls.Edit
Public txtBuffsProfile As DecalControls.Edit
Public txtLootProfile As DecalControls.Edit
Public WithEvents cmdSaveMacroProfile As DecalControls.PushButton
Attribute cmdSaveMacroProfile.VB_VarHelpID = -1
Public WithEvents cmdSaveBuffsProfile As DecalControls.PushButton
Attribute cmdSaveBuffsProfile.VB_VarHelpID = -1
Public WithEvents cmdSaveLootProfile As DecalControls.PushButton
Attribute cmdSaveLootProfile.VB_VarHelpID = -1

Public WithEvents btnSetSell As DecalControls.PushButton
Attribute btnSetSell.VB_VarHelpID = -1
Public WithEvents btnSetBuy As DecalControls.PushButton
Attribute btnSetBuy.VB_VarHelpID = -1
Public WithEvents btnStartSell As DecalControls.PushButton
Attribute btnStartSell.VB_VarHelpID = -1
Public chkSellAll As DecalControls.CheckBox

Public WithEvents btnRebuff As DecalControls.PushButton
Attribute btnRebuff.VB_VarHelpID = -1
Public WithEvents btnSendXpHour As DecalControls.PushButton
Attribute btnSendXpHour.VB_VarHelpID = -1
Public WithEvents btnSendComps As DecalControls.PushButton
Attribute btnSendComps.VB_VarHelpID = -1
Public WithEvents btnDebug As DecalControls.PushButton
Attribute btnDebug.VB_VarHelpID = -1

Public lblElapsed As DecalControls.StaticText
Public lblNextRebuff As DecalControls.StaticText
Public lblTapersHour As DecalControls.StaticText
Public lblTotalXp As DecalControls.StaticText
Public lblXpEarned As DecalControls.StaticText
Public lblXpHour As DecalControls.StaticText
Public lblXpMinute As DecalControls.StaticText
Public lblNextLevel As DecalControls.StaticText

'Info tab
Public WithEvents btnResetXp As DecalControls.PushButton
Attribute btnResetXp.VB_VarHelpID = -1
Public progBuffs As DecalControls.Progress

Private m_bLoaded As Boolean

'Controls declaration
Public Function Init(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler

    'Set interface module name
    InterfaceName = "UIMain"
    
    MyDebug InterfaceName & ".Init() -- Begin"
    
    'Decal Controls initialisation
1    Set chkEnable = g_MainView.Control("chkEnable")
    
    Set chMacroProfile = g_MainView.Control("chMacroProfile")
    Set chBuffsProfile = g_MainView.Control("chBuffsProfile")
    Set chLootProfile = g_MainView.Control("chLootProfile")
    Set txtMacroProfile = g_MainView.Control("txtMacroProfile")
    Set txtBuffsProfile = g_MainView.Control("txtBuffsProfile")
    Set txtLootProfile = g_MainView.Control("txtLootProfile")
    Set cmdSaveMacroProfile = g_MainView.Control("cmdSaveMacroProfile")
    Set cmdSaveBuffsProfile = g_MainView.Control("cmdSaveBuffsProfile")
    Set cmdSaveLootProfile = g_MainView.Control("cmdSaveLootProfile")

3   Set btnSetSell = g_MainView.Control("btnSetSell")
    Set btnSetBuy = g_MainView.Control("btnSetBuy")
    Set btnStartSell = g_MainView.Control("btnStartSell")
    Set chkSellAll = g_MainView.Control("chkSellAll")

4   Set chCombatType = g_MainView.Control("chCombatType")
    Set btnPause = g_MainView.Control("btnPause")
    Set btnRebuff = g_MainView.Control("btnRebuff")
    Set btnSendXpHour = g_MainView.Control("btnSendXpHour")
    Set btnSendComps = g_MainView.Control("btnSendComps")
    Set btnDebug = g_MainView.Control("btnDebug")
   
5   Set lblElapsed = g_MainView.Control("lblElapsed")
6   Set lblNextRebuff = g_MainView.Control("lblNextRebuff")
7   Set lblTapersHour = g_MainView.Control("lblTapersHour")
8   Set lblTotalXp = g_MainView.Control("lblTotalXp")
9   Set lblXpEarned = g_MainView.Control("lblXpEarned")
10  Set lblXpHour = g_MainView.Control("lblXpHour")
11  Set lblXpMinute = g_MainView.Control("lblXpMinute")
12  Set lblNextLevel = g_MainView.Control("lblNextLevel")
   
20  Set lblStatus = g_MainView.Control("lblStatus")
    Set btnResetXp = g_MainView.Control("btnResetXp")
    
21  Set progBuffs = g_MainView.Control("progBuffs")
    Set stVersion = g_MainView.Control("stVersion")
    
22  Set lblBuyTarget = g_MainView.Control("lblBuyTarget")
    Set lblSellTarget = g_MainView.Control("lblSellTarget")
    
    'Default Control Initialization
23  chkEnable.Checked = False
    progBuffs.DecalDrawText = False
    progBuffs.PreText = "Buffs : "
    chkSellAll.Checked = True
    
30    stVersion.Text = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
    
    Init = True
    MyDebug InterfaceName & ".Init() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Init = False
    PrintErrorMessage "interface:" & InterfaceName & " Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Function

Public Function LoadControlsValue(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler

    LoadControlsValue = False
    
    '***************************************************************************
    chkEnable.Checked = False
    chCombatType.Selected = g_Settings.Profile.MacroCfg.GetValue("chCombatType", 0)
    
    'Puts the different buff profiles available with the current profile in the Buffs Profile choice list
    'Call UpdateProfileList
    Call UpdateStatus
    
    If Not m_bLoaded Then
        m_bLoaded = True
        Call UpdateMacroProfileList
        Call UpdateBuffsProfileList
        Call UpdateLootProfileList
    End If
    '***************************************************************************
    
    LoadControlsValue = True
    
Fin:
    Exit Function
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage "(" & InterfaceName & " LoadControlsValue) - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    
    MyDebug "[" & InterfaceName & "] Saving controls settings"
    
    With g_Settings.Profile.MacroCfg
        .SaveChoice chCombatType, "chCombatType"
    End With
    
    bRet = True
Fin:
    SaveControlSettings = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage InterfaceName & ".SaveControlSettings - " & Err.Description
    Resume Fin
End Function

Public Function Unload() As Boolean
On Error GoTo Error_Handler
    
    MyDebug InterfaceName & ".Unload() -- Begin"
    
    Set chkEnable = Nothing
    
    Set btnPause = Nothing
    Set stVersion = Nothing
    Set lblStatus = Nothing
    Set lblBuyTarget = Nothing
    Set lblSellTarget = Nothing

    Set chMacroProfile = Nothing
    Set chBuffsProfile = Nothing
    Set chLootProfile = Nothing
    Set txtMacroProfile = Nothing
    Set txtBuffsProfile = Nothing
    Set txtLootProfile = Nothing
    Set cmdSaveMacroProfile = Nothing
    Set cmdSaveBuffsProfile = Nothing
    Set cmdSaveLootProfile = Nothing
    
    Set btnSetSell = Nothing
    Set btnSetBuy = Nothing
    Set btnStartSell = Nothing
    Set chkSellAll = Nothing

    Set lblElapsed = Nothing
    Set lblNextRebuff = Nothing
    Set lblTapersHour = Nothing
    Set lblTotalXp = Nothing
    Set lblXpEarned = Nothing
    Set lblXpHour = Nothing
    Set lblXpMinute = Nothing
    Set lblNextLevel = Nothing
    
    Set btnRebuff = Nothing
    Set btnResetXp = Nothing
    Set btnSendXpHour = Nothing
    Set btnSendComps = Nothing
    Set btnDebug = Nothing
    
    Set chCombatType = Nothing
    Set progBuffs = Nothing
    
    Unload = True
    MyDebug InterfaceName & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage "(interface:" & InterfaceName & ") Unload - " & Err.Description
    Resume Fin

End Function

Private Sub btnDebug_Accepted(ByVal nID As Long)
    Call g_Data.SetDebugMode(Not g_Data.mDebugMode)
    If g_Data.mDebugMode Then
        PrintMessage "Debug Mode : ON"
    Else
        PrintMessage "Debug Mode : OFF"
    End If
End Sub

Private Sub btnPause_Accepted(ByVal nID As Long)
    Call TogglePause
End Sub

Private Sub btnRebuff_Accepted(ByVal nID As Long)
    If Valid(g_Macro) Then Call g_Macro.ForceRebuff
End Sub

Private Sub btnResetXp_Accepted(ByVal nID As Long)
    If Valid(g_Macro) Then
        Call g_Macro.ResetXpStats
        Call g_DOT.Reset
    End If
End Sub


Private Sub btnSendComps_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler
    Dim sMsg As String
    
    sMsg = g_RemoteCmd.CmdCompsString

    PrintMessage sMsg

    If (Not g_ui.Options.chkReportToIrcOnly.Checked) And (g_Objects.Fellowship.NumMembers > 0) Then
        Call SendFellowshipMessage(sMsg)
    End If

    If g_ui.Irc.ConnectedToChannel Then
        Call g_ui.Irc.SendChanMessage(sMsg)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "btnSendComps_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub btnSendXpHour_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler
    Dim sMsg As String
    
    sMsg = g_RemoteCmd.CmdReportString
    
    PrintMessage sMsg

    If (Not g_ui.Options.chkReportToIrcOnly.Checked) And (g_Objects.Fellowship.NumMembers > 0) Then
        Call SendFellowshipMessage(sMsg)
    End If

    If g_ui.Irc.ConnectedToChannel Then
        Call g_ui.Irc.SendChanMessage(sMsg)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "btnSendXpHour_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub chCombatType_Change(ByVal nID As Long, ByVal nIndex As Long)
    If Valid(g_Macro) And nIndex >= 0 Then
        g_Macro.CombatType = chCombatType.Selected
    End If
End Sub

Private Sub chkEnable_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    Call ToggleMacro(bChecked)
End Sub

Private Sub Class_Initialize()
    m_bLoaded = False
End Sub

Private Sub Class_Terminate()
    Call Unload
End Sub

Public Sub ResetBuffsProgressBar()
    progBuffs.MaxValue = 0
    progBuffs.Value = 0
    progBuffs.DecalDrawText = False
End Sub

Public Sub UpdateStats()
On Error GoTo ErrorHandler

    'Update other stats fields
    lblTotalXp.Text = g_Macro.TotalXpString
    lblXpEarned.Text = g_Macro.XpEarnedString
    lblXpHour.Text = g_Macro.XpHourString
    lblXpMinute.Text = g_Macro.XpMinuteString
    lblNextLevel.Text = g_Macro.TimeUntilNextLevelString
    lblTapersHour.Text = g_Macro.TapersHourString

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "UpdateStats - " & Err.Description
    Resume Fin
End Sub
Public Sub UpdateStatus()
On Error GoTo ErrorHandler

    Dim sOut As String
    sOut = "Off"
    
    If g_Macro.Active Then
        sOut = g_Macro.GetStateString
        If g_Spells.Casting Then sOut = sOut & " (casting)"
        If IsBusy(TurboMode) Then sOut = sOut & " <!>"
    End If
    
    If g_Macro.Paused Then sOut = "Paused"
    
    lblStatus.Text = sOut
    lblElapsed.Text = g_Macro.ElapsedTimeString
    lblNextRebuff.Text = g_Macro.NextRebuffTimeString

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "UpdateStatus - " & Err.Description
    Resume Fin
End Sub

Private Sub UpdateProfileList(ByRef chControl As DecalControls.Choice, ByVal sBaseFolder As String, ByVal sRequiredFile As String, ByVal sCurrent As String)
On Error GoTo ErrorHandler

    Dim oFile As clsFileInfo
    Dim oMainFolder As Scripting.Folder
    Dim oCurFolder As Scripting.Folder
    Dim oFS As New Scripting.FileSystemObject
    Dim i As Integer
    
    MyDebug "UpdateProfileList"
    
    Call chControl.Clear
    i = -1
    
    Set oMainFolder = oFS.GetFolder(sBaseFolder)
    
    If Valid(oMainFolder) Then
        For Each oCurFolder In oMainFolder.SubFolders
            If oFS.FileExists(oCurFolder.Path & "\" & sRequiredFile) Then
                Call chControl.AddChoice(oCurFolder.Name)
                If SameText(oCurFolder.Name, sCurrent) Then
                    i = chControl.ChoiceCount - 1
                End If
            End If
        Next oCurFolder
    End If
    
    If (i <> -1) And (chControl.ChoiceCount > i) Then
       chControl.Selected = i
    End If
    
Fin:
    Set oFile = Nothing
    Set oCurFolder = Nothing
    Set oMainFolder = Nothing
    Set oFS = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.UpdateProfileList(" & sBaseFolder & ") - " & Err.Description & " (line " & Erl & ")"
    Resume Fin
End Sub

Private Sub UpdateBuffsProfileList()
On Error GoTo ErrorHandler

    Call UpdateProfileList(chBuffsProfile, g_Settings.Profile.FullPath & "\" & FOLDER_PROFILE_BUFFS, FILE_BUFFS_CONFIG, g_Settings.Profile.BuffsProfileName)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.UpdateBuffsProfileList - " & Err.Description
    Resume Fin
End Sub


Private Sub UpdateMacroProfileList()
On Error GoTo ErrorHandler
    
    Call UpdateProfileList(chMacroProfile, g_Settings.Profile.FullPath & "\" & FOLDER_PROFILE_MACRO, FILE_MACRO_CONFIG, g_Settings.Profile.MacroProfileName)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.UpdateMacroProfileList"
    Resume Fin
End Sub

Private Sub UpdateLootProfileList()
On Error GoTo ErrorHandler

    Call UpdateProfileList(chLootProfile, g_Settings.GetDataFolder & "\" & PATH_PROFILE_LOOT, FILE_LOOT_CONFIG, g_Settings.Profile.LootProfileName)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.UpdateLootProfileList - " & Err.Description
    Resume Fin
End Sub

Private Sub MakeProfilePath(ByVal sBasePath As String, ByVal sProfile As String)
On Error GoTo ErrorHandler

    Dim oFS As New Scripting.FileSystemObject
    Dim sProfilePath As String
    sProfilePath = sBasePath & "\" & sProfile
    
    If Not oFS.FolderExists(sProfilePath) Then
        Call oFS.CreateFolder(sProfilePath)
    End If
    
Fin:
    Set oFS = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.MakeProfilePath - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdSaveBuffsProfile_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    
    Dim sProfile As String
    sProfile = Trim(txtBuffsProfile.Text)
    
    If sProfile <> "" Then
        Call MakeProfilePath(g_Settings.Profile.FullPath & "\" & FOLDER_PROFILE_BUFFS, sProfile)
        Call g_Settings.Profile.SetBuffsProfile(sProfile)
        Call g_Settings.SavePluginConfiguration
        Call UpdateBuffsProfileList
    Else
        PrintMessage "Please enter a valid profile name."
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.cmdSaveBuffsProfile_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdSaveLootProfile_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    
    Dim sProfile As String
    sProfile = Trim(txtLootProfile.Text)
    
    If sProfile <> "" Then
        Call MakeProfilePath(g_Settings.GetDataFolder & "\" & PATH_PROFILE_LOOT, sProfile)
        Call g_Settings.Profile.SetLootProfile(sProfile)
        Call g_Settings.SavePluginConfiguration
        Call UpdateLootProfileList
    Else
        PrintMessage "Please enter a valid profile name."
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.cmdSaveLootProfile_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdSaveMacroProfile_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    
    Dim sProfile As String
    sProfile = Trim(txtMacroProfile.Text)
    
    If sProfile <> "" Then
        Call MakeProfilePath(g_Settings.Profile.FullPath & "\" & FOLDER_PROFILE_MACRO, sProfile)
        Call g_Settings.Profile.SetMacroProfile(sProfile)
        Call g_Settings.SavePluginConfiguration
        Call UpdateMacroProfileList
    Else
        PrintMessage "Please enter a valid profile name."
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.cmdSaveMacroProfile_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub btnSetBuy_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    If g_Hooks.VendorID = 0 Then
        PrintMessage "You must open a Vendor first!"
        GoTo Fin
    End If
    
    If Not (g_Hooks.CurrentSelection <> 0) Then
        MyDebug "btnSetBuy: g_Hooks.CurrentSelection not valid"
        PrintMessage "Please select a valid object"
    Else
        MyDebug "SetBuy_Accepted: g_hooks.cursel: " & g_Hooks.CurrentSelection
        Dim obj As acObject
        'Set obj = g_Objects.FindObject(g_Hooks.CurrentSelection)
        'Set obj = g_ds.GameObjects.Selected
        
        Set obj = PhatLoot.findInVendorInv(g_Hooks.CurrentSelection)
        
        If Valid(obj) Then
            MyDebug "btnSetBuy, found valid obj: " & obj.Name & " value: " & obj.Value
            Call PhatLoot.setBuyTarget(obj)
            lblBuyTarget.Text = obj.Name
            PrintMessage "Set buy target as: " & obj.Name
        Else
            MyDebug "btnSetBuy: could not find obj with PhatLoot.findInVendorInv"
            PrintMessage "First select the object to buy!"
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.btnSetBuy_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub btnSetSell_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    If g_Hooks.VendorID = 0 Then
        PrintMessage "You must open a Vendor window first!"
        GoTo Fin
    End If
    
    If Not (g_Hooks.CurrentSelection <> 0) Then
        MyDebug "btnSetSell: g_Hooks.CurrentlSelection not valid"
        PrintMessage "Please select the object to sell"
    Else
        Dim obj As acObject
        'Set obj = g_ds.GameObjects.Selected
        Set obj = g_Objects.FindObject(g_Hooks.CurrentSelection)
        If Valid(obj) Then
            MyDebug "btnSetSell found valid obj: " & obj.Name
            Call PhatLoot.setSellTarget(obj)
            lblSellTarget.Text = obj.Name
            PrintMessage "Set sell target as: " & obj.Name
        Else
            MyDebug "btnSelSell: not find obj in g_ds.GameObjects.Selected"
            PrintMessage "First select the object to Sell!"
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.btnSetSell_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub btnStartSell_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    If Not g_Macro.Active Then
        PrintErrorMessage "LT must be running to buy/sell"
        GoTo Fin
    End If
    
    If g_Hooks.VendorID = 0 Then
        PrintMessage "You must open a Vendor window first!"
        GoTo Fin
    End If
    
    If g_Macro.State = ST_BUYSELL Then
        PrintMessage "Stop Selling"
        Call PhatLoot.StopSell
        Exit Sub
    End If
    
    If PhatLoot.IsReadyToSell Then
        PrintMessage "Starting!"
        Call PhatLoot.StartSell
    Else
        PrintMessage "Please select items to buy/sell first!"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMain.btnStartSell_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub chBuffsProfile_Change(ByVal nID As Long, ByVal nIndex As Long)
    If nIndex >= 0 And g_bProfileLoaded Then
        Dim sProfile As String
        sProfile = chBuffsProfile.Text(nIndex)
        
        'Save current profile
        Call g_Settings.SavePluginConfiguration
        
        'Change the profile
        Call g_Settings.Profile.LoadBuffsProfile(sProfile)
        PrintMessage "Using Buff Profile : " & sProfile
        
        'Load the controls values
        Call g_ui.Buffs.LoadControlsValue
    End If
End Sub

Private Sub chMacroProfile_Change(ByVal nID As Long, ByVal nIndex As Long)
    If nIndex >= 0 And g_bProfileLoaded Then
        Dim sProfile As String
        sProfile = chMacroProfile.Text(nIndex)
        
        'Save current profile
        Call g_Settings.SavePluginConfiguration
        
        'Change the profile
        Call g_Settings.Profile.LoadMacroProfile(sProfile)
        PrintMessage "Using Macro Profile : " & sProfile
        
        'Load the controls values
        Call LoadControlsValue          'load g_ui.main
        Call g_ui.Macro.LoadControlsValue
        Call g_ui.Options.LoadControlsValue
    End If
End Sub

Private Sub chLootProfile_Change(ByVal nID As Long, ByVal nIndex As Long)
    If nIndex >= 0 And g_bProfileLoaded Then
        Dim sProfile As String
        sProfile = chLootProfile.Text(nIndex)
        
        'Save current profile
        Call g_Settings.SavePluginConfiguration
        
        'Change the profile
        Call g_Settings.Profile.LoadLootProfile(sProfile)
        PrintMessage "Using Loot Profile : " & sProfile
        
        'Load the controls values
        Call g_ui.Loot.LoadControlsValue
        
    End If
End Sub
