VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================
' Macro Options Controls
'===========================================
Option Explicit

Private Enum ePickupListColumns
    COL_ENABLED
    COL_ITEM_NAME
    COL_EXACT_MATCH
End Enum

Private Const INTERFACE_NAME = "UIOptions"

'Filters (Options tab)
Public chkFilterMelee As DecalControls.CheckBox
Attribute chkFilterMelee.VB_VarHelpID = -1
Public chkFilterMeleeEvade As DecalControls.CheckBox
Public chkFilterSpellcasting As DecalControls.CheckBox
Attribute chkFilterSpellcasting.VB_VarHelpID = -1
Public chkFilterDeathMsg As DecalControls.CheckBox
Attribute chkFilterDeathMsg.VB_VarHelpID = -1
Public chkFilterGlobalChat As DecalControls.CheckBox
Public chkFilterGlobalTrade As DecalControls.CheckBox
Public chkFilterGlobalLFG As DecalControls.CheckBox

' LTx generated messages
Public chkFilterLTMsg As DecalControls.CheckBox
Attribute chkFilterLTMsg.VB_VarHelpID = -1
Public chkFilterLootMsg As DecalControls.CheckBox
Public chkFilterSalvageMsg As DecalControls.CheckBox
Public chkFilterIdMsg As DecalControls.CheckBox

'Alerts
Public chkAlertAdmin As DecalControls.CheckBox
Public chkAlertTell As DecalControls.CheckBox
Public chkAlertUnfriendly As DecalControls.CheckBox
Public chkAlertFellowDead As DecalControls.CheckBox
Public chkAlertOpenChat As DecalControls.CheckBox
Public chkAlertDeath As DecalControls.CheckBox
Public chkAlertRare As DecalControls.CheckBox

'Coms
Public chkAnswerChanCmd As DecalControls.CheckBox
Public chkCmdWhere As DecalControls.CheckBox
Public txtLocation As DecalControls.Edit
Public chkReportToIrcOnly As DecalControls.CheckBox
Public chkAnswerIngameCmd As DecalControls.CheckBox
 
'Options>HUD
Public WithEvents chkEnableHUD As DecalControls.CheckBox
Attribute chkEnableHUD.VB_VarHelpID = -1
Public WithEvents chkHUDMacro As DecalControls.CheckBox
Attribute chkHUDMacro.VB_VarHelpID = -1
Public WithEvents chkHUDStats As DecalControls.CheckBox
Attribute chkHUDStats.VB_VarHelpID = -1
Public WithEvents chkHUDInventory As DecalControls.CheckBox
Attribute chkHUDInventory.VB_VarHelpID = -1
Public WithEvents chkHUDDOT As DecalControls.CheckBox
Attribute chkHUDDOT.VB_VarHelpID = -1
Public WithEvents chkHUDDebug As DecalControls.CheckBox
Attribute chkHUDDebug.VB_VarHelpID = -1
Public WithEvents chkLockHuds As DecalControls.CheckBox
Attribute chkLockHuds.VB_VarHelpID = -1
Public chk3DTarget As DecalControls.CheckBox

'Friend List
Public WithEvents lstFriends As DecalControls.list
Attribute lstFriends.VB_VarHelpID = -1
Public WithEvents chkOnlyFriendsList As DecalControls.CheckBox
Attribute chkOnlyFriendsList.VB_VarHelpID = -1
Public txtAddFriend As DecalControls.Edit
Public WithEvents btnAddFriend As DecalControls.PushButton
Attribute btnAddFriend.VB_VarHelpID = -1

'Extra
Public chkDisableLogs As DecalControls.CheckBox

'Detect
Public WithEvents lstDetect As DecalControls.list
Attribute lstDetect.VB_VarHelpID = -1
Public WithEvents chkDetectList As DecalControls.CheckBox
Attribute chkDetectList.VB_VarHelpID = -1
Public txtAddDetect As DecalControls.Edit
Public WithEvents btnAddDetect As DecalControls.PushButton
Attribute btnAddDetect.VB_VarHelpID = -1
Public WithEvents btnSetDetect As DecalControls.PushButton
Attribute btnSetDetect.VB_VarHelpID = -1

'Hud locations
Public macroHUDposX As Integer
Public macroHUDposY As Integer
Public statsHUDposX As Integer
Public statsHUDposY As Integer
Public invHUDposX As Integer
Public invHUDposY As Integer
Public debugHUDposX As Integer
Public debugHUDposY As Integer
Public dotHUDposX As Integer
Public dotHUDposY As Integer


Private Sub Class_Terminate()
    Call Unload
End Sub

Public Function Init() As Boolean
On Error GoTo Error_Handler

    'Initialize
    Init = False
    
    MyDebug INTERFACE_NAME & ".Init() -- Begin"
    
    'Create Controls
    Set chkFilterMelee = g_MainView.Control("chkFilterMelee")
    Set chkFilterMeleeEvade = g_MainView.Control("chkFilterMeleeEvade")
    Set chkFilterSpellcasting = g_MainView.Control("chkFilterSpellcasting")
    Set chkFilterDeathMsg = g_MainView.Control("chkFilterDeathMsg")
    Set chkFilterGlobalChat = g_MainView.Control("chkFilterGlobalChat")
    Set chkFilterGlobalTrade = g_MainView.Control("chkFilterGlobalTrade")
    Set chkFilterGlobalLFG = g_MainView.Control("chkFilterGlobalLFG")
    
    Set chkFilterLTMsg = g_MainView.Control("chkFilterLTMsg")
    Set chkFilterLootMsg = g_MainView.Control("chkFilterLootMsg")
    Set chkFilterSalvageMsg = g_MainView.Control("chkFilterSalvageMsg")
    Set chkFilterIdMsg = g_MainView.Control("chkFilterIdMsg")
    
    Set chkAlertAdmin = g_MainView.Control("chkAlertAdmin")
    Set chkAlertTell = g_MainView.Control("chkAlertTell")
    Set chkAlertUnfriendly = g_MainView.Control("chkAlertUnfriendly")
    Set chkAlertFellowDead = g_MainView.Control("chkAlertFellowDead")
    Set chkAlertOpenChat = g_MainView.Control("chkAlertOpenChat")
    Set chkAlertDeath = g_MainView.Control("chkAlertDeath")
    Set chkAlertRare = g_MainView.Control("chkAlertRare")
    
    Set chkAnswerChanCmd = g_MainView.Control("chkAnswerChanCmd")
    Set chkCmdWhere = g_MainView.Control("chkCmdWhere")
    Set txtLocation = g_MainView.Control("txtLocation")
    Set chkReportToIrcOnly = g_MainView.Control("chkReportToIrcOnly")
    Set chkAnswerIngameCmd = g_MainView.Control("chkAnswerIngameCmd")
    
    Set chkEnableHUD = g_MainView.Control("chkEnableHUD")
    Set chkHUDMacro = g_MainView.Control("chkHUDMacro")
    Set chkHUDStats = g_MainView.Control("chkHUDStats")
    Set chkHUDInventory = g_MainView.Control("chkHUDInventory")
    Set chkHUDDebug = g_MainView.Control("chkHUDDebug")
    Set chkHUDDOT = g_MainView.Control("chkHUDDOT")
    Set chkLockHuds = g_MainView.Control("chkLockHuds")
    Set chk3DTarget = g_MainView.Control("chk3DTarget")
    
    Set lstFriends = g_MainView.Control("lstFriends")
    Set btnAddFriend = g_MainView.Control("btnAddFriend")
    Set txtAddFriend = g_MainView.Control("txtAddFriend")
    Set chkOnlyFriendsList = g_MainView.Control("chkOnlyFriendsList")
    
    Set chkDisableLogs = g_MainView.Control("chkDisableLogs")
    
    Set lstDetect = g_MainView.Control("lstDetect")
    Set btnAddDetect = g_MainView.Control("btnAddDetect")
    Set btnSetDetect = g_MainView.Control("btnSetDetect")
    Set txtAddDetect = g_MainView.Control("txtAddDetect")
    Set chkDetectList = g_MainView.Control("chkDetectList")
    
    Init = True
    MyDebug INTERFACE_NAME & ".Init() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Init = False
    PrintErrorMessage INTERFACE_NAME & ".Init - " & Err.Description
    Resume Fin
End Function

Public Function LoadControlsValue() As Boolean
On Error GoTo Error_Handler
    
    MyDebug "[" & INTERFACE_NAME & "] Loading controls settings"
    
    '***************************************************************************
    
    chkFilterMelee.Checked = g_Settings.GetValue("chkFilterMelee", False)
    chkFilterMeleeEvade.Checked = g_Settings.GetValue("chkFilterMeleeEvade", False)
    chkFilterSpellcasting.Checked = g_Settings.GetValue("chkFilterSpellcasting", False)
    chkFilterDeathMsg.Checked = g_Settings.GetValue("chkFilterDeathMsg", False)
    chkFilterGlobalChat.Checked = g_Settings.GetValue("chkFilterGlobalChat", False)
    chkFilterGlobalTrade.Checked = g_Settings.GetValue("chkFilterGlobalTrade", False)
    chkFilterGlobalLFG.Checked = g_Settings.GetValue("chkFilterGlobalLFG", False)
    
    chkFilterLTMsg.Checked = g_Settings.GetValue("chkFilterLTMsg", False)
    chkFilterLootMsg.Checked = g_Settings.GetValue("chkFilterLootMsg", False)
    chkFilterSalvageMsg.Checked = g_Settings.GetValue("chkFilterSalvageMsg", False)
    chkFilterIdMsg.Checked = g_Settings.GetValue("chkFilterIdMsg", False)
    
    chkAlertAdmin.Checked = g_Settings.GetValue("chkAlertAdmin", True)
    chkAlertTell.Checked = g_Settings.GetValue("chkAlertTell", False)
    chkAlertUnfriendly.Checked = g_Settings.GetValue("chkAlertUnfriendly", False)
    chkAlertFellowDead.Checked = g_Settings.GetValue("chkAlertFellowDead", False)
    chkAlertOpenChat.Checked = g_Settings.GetValue("chkAlertOpenChat", False)
    chkAlertDeath.Checked = g_Settings.GetValue("chkAlertDeath", False)
    chkAlertRare.Checked = g_Settings.GetValue("chkAlertRare", True)
    
    chkAnswerChanCmd.Checked = g_Settings.GetValue("chkAnswerChanCmd", True)
    chkReportToIrcOnly.Checked = g_Settings.GetValue("chkReportToIrcOnly", False)
    chkAnswerIngameCmd.Checked = g_Settings.GetValue("chkAnswerIngameCmd", False)
    
    chkEnableHUD.Checked = g_Settings.GetValue("chkEnableHUD", True)
    chkHUDMacro.Checked = g_Settings.GetValue("chkHUDMacro", True)
    chkHUDStats.Checked = g_Settings.GetValue("chkHUDStats", True)
    chkHUDInventory.Checked = g_Settings.GetValue("chkHUDInventory", True)
    chkHUDDebug.Checked = g_Settings.GetValue("chkHUDDebug", False)
    chkHUDDOT.Checked = g_Settings.GetValue("chkHUDDOT", False)
    chkLockHuds.Checked = g_Settings.GetValue("chkLockHuds", True)
    chk3DTarget.Checked = g_Settings.GetValue("chk3DTarget", True)
    
    chkOnlyFriendsList.Checked = g_Settings.GetValue("chkOnlyFriendsList", False)
    
    chkDisableLogs.Checked = g_Settings.GetValue("chkDisableLogs", False)
    
    chkDetectList.Checked = g_Settings.GetValue("chkDetectList", False)
    
    '***************************************************************************
    
    chkCmdWhere.Checked = g_Settings.Profile.MacroCfg.GetValue("chkCmdWhere", True)
    txtLocation.Text = g_Settings.Profile.MacroCfg.GetValue("txtLocation", "You wished you knew!")
    
    '***************************************************************************
    
    Call g_ui.Loot.LoadItemsList(lstFriends, GetFriendsListPath)
    Call g_ui.Loot.LoadItemsList(lstDetect, GetDetectListPath)
    
    LoadControlsValue = True
    
Fin:
    Exit Function
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage INTERFACE_NAME & ".LoadControlsValue - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    
    MyDebug "[" & INTERFACE_NAME & "] Saving controls settings"
    
    With g_Settings.Profile.Config
        .SaveCheckbox chkFilterMelee, "chkFilterMelee"
        .SaveCheckbox chkFilterMeleeEvade, "chkFilterMeleeEvade"
        .SaveCheckbox chkFilterSpellcasting, "chkFilterSpellcasting"
        .SaveCheckbox chkFilterDeathMsg, "chkFilterDeathMsg"
        .SaveCheckbox chkFilterGlobalChat, "chkFilterGlobalChat"
        .SaveCheckbox chkFilterGlobalTrade, "chkFilterGlobalTrade"
        .SaveCheckbox chkFilterGlobalLFG, "chkFilterGlobalLFG"
        
        .SaveCheckbox chkFilterLTMsg, "chkFilterLTMsg"
        .SaveCheckbox chkFilterLootMsg, "chkFilterLootMsg"
        .SaveCheckbox chkFilterSalvageMsg, "chkFilterSalvageMsg"
        .SaveCheckbox chkFilterIdMsg, "chkFilterIdMsg"
        
        .SaveCheckbox chkAlertAdmin, "chkAlertAdmin"
        .SaveCheckbox chkAlertTell, "chkAlertTell"
        .SaveCheckbox chkAlertUnfriendly, "chkAlertUnfriendly"
        .SaveCheckbox chkAlertFellowDead, "chkAlertFellowDead"
        .SaveCheckbox chkAlertOpenChat, "chkAlertOpenChat"
        .SaveCheckbox chkAlertDeath, "chkAlertDeath"
        .SaveCheckbox chkAlertRare, "chkAlertRare"
        
        .SaveCheckbox chkAnswerIngameCmd, "chkAnswerIngameCmd"
        .SaveCheckbox chkAnswerChanCmd, "chkAnswerChanCmd"
        .SaveCheckbox chkReportToIrcOnly, "chkReportToIrcOnly"
        
        .SaveCheckbox chkEnableHUD, "chkEnableHUD"
        .SaveCheckbox chkHUDMacro, "chkHUDMacro"
        .SaveCheckbox chkHUDStats, "chkHUDStats"
        .SaveCheckbox chkHUDInventory, "chkHUDInventory"
        .SaveCheckbox chkHUDDebug, "chkHUDDebug"
        .SaveCheckbox chkHUDDOT, "chkHUDDOT"
        .SaveCheckbox chk3DTarget, "chk3DTarget"
        
        .SaveCheckbox chkOnlyFriendsList, "chkOnlyFriendsList"
        
        .SaveCheckbox chkDetectList, "chkDetectList"

    End With
    
    Call g_ui.Loot.SaveItemsList(lstFriends, GetFriendsListPath)
    Call g_ui.Loot.SaveItemsList(lstDetect, GetDetectListPath)
        
    With g_Settings.Profile.MacroCfg
        .SaveCheckbox chkCmdWhere, "chkCmdWhere"
        .SaveTextbox txtLocation, "txtLocation"
    End With
    
    bRet = True
Fin:
    SaveControlSettings = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage INTERFACE_NAME & ".SaveControlSettings - " & Err.Description
    Resume Fin
End Function

Public Function Unload() As Boolean
On Error GoTo Error_Handler
    
    MyDebug INTERFACE_NAME & ".Unload() -- Begin"
    'Unload initialization
    Unload = False

    Set chkFilterMelee = Nothing
    Set chkFilterMeleeEvade = Nothing
    Set chkFilterSpellcasting = Nothing
    Set chkFilterDeathMsg = Nothing
    Set chkFilterGlobalChat = Nothing
    Set chkFilterGlobalTrade = Nothing
    Set chkFilterGlobalLFG = Nothing
    Set chkFilterLTMsg = Nothing
    Set chkFilterLootMsg = Nothing
    Set chkFilterSalvageMsg = Nothing
    Set chkFilterIdMsg = Nothing
    
    Set chkAlertAdmin = Nothing
    Set chkAlertTell = Nothing
    Set chkAlertUnfriendly = Nothing
    Set chkAlertFellowDead = Nothing
    Set chkAlertOpenChat = Nothing
    Set chkAlertDeath = Nothing
    Set chkAlertRare = Nothing
    
    Set chkAnswerIngameCmd = Nothing
    Set chkAnswerChanCmd = Nothing
    Set chkCmdWhere = Nothing
    Set txtLocation = Nothing
    Set chkReportToIrcOnly = Nothing
    
    Set chkEnableHUD = Nothing
    Set chkHUDMacro = Nothing
    Set chkHUDStats = Nothing
    Set chkHUDInventory = Nothing
    Set chkHUDDebug = Nothing
    Set chkHUDDOT = Nothing
    Set chkLockHuds = Nothing
    Set chk3DTarget = Nothing
    
    Set lstFriends = Nothing
    Set btnAddFriend = Nothing
    Set txtAddFriend = Nothing
    Set chkOnlyFriendsList = Nothing
    
    Set chkDisableLogs = Nothing
    
    Set lstDetect = Nothing
    Set btnAddDetect = Nothing
    Set txtAddDetect = Nothing
    Set btnSetDetect = Nothing
    Set chkDetectList = Nothing
    
    'Unload complete
    Unload = True
    MyDebug INTERFACE_NAME & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage INTERFACE_NAME & ".Unload - " & Err.Description
    Resume Fin
End Function

'---------------------------
' Functions
'---------------------------

Private Function GetFriendsListPath() As String
    GetFriendsListPath = g_Settings.Profile.FullPath & "\" & FILE_ANTIBAN_FRIENDS_LIST
End Function

Public Function NameInFriendsList(aName As String) As Boolean
    NameInFriendsList = g_ui.Loot.ItemCheckedInList(lstFriends, aName)
End Function

Private Function GetDetectListPath() As String
    GetDetectListPath = g_Settings.Profile.FullPath & "\" & FILE_DETECT_LIST
End Function

Public Function NameInDetectList(aName As String, isCritter As Boolean) As Boolean
    NameInDetectList = findDetectMatchInList(lstDetect, aName, isCritter)
End Function

'returns False if couldn't find item
Private Function findDetectMatchInList(ByRef lstControl As DecalControls.list, ByVal FullItemName As String, ByVal isCritter As Boolean) As Boolean
On Error GoTo Error_Handler

    Dim i As Integer
    Dim isEnabled As Boolean
    Dim curItemName As String
    Dim critterOnly As Boolean
    Dim bRet As Boolean

    'default return value
    bRet = False
    
    For i = 0 To lstControl.Count - 1
        isEnabled = lstControl.Data(COL_ENABLED, i)
        curItemName = lstControl.Data(COL_ITEM_NAME, i)
        critterOnly = lstControl.Data(COL_EXACT_MATCH, i)
        
        If isEnabled Then
            If isCritter Then
                'Only try entries that are marked "Critter"
                If critterOnly And InStr(LCase(FullItemName), LCase(curItemName)) Then
                    bRet = True
                End If
            Else
                If (Not critterOnly) And InStr(LCase(FullItemName), LCase(curItemName)) Then
                    bRet = True
                End If
            End If
        End If
        
    Next i
    
Fin:
    findDetectMatchInList = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage "clsUIOptions.findCritterMatchInList: Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Function


'---------------------------
' Events
'---------------------------

Private Sub btnAddFriend_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    If g_ui.Loot.AddItemToList(lstFriends, txtAddFriend.Text, True, False) Then
        Call g_ui.Loot.SaveItemsList(lstFriends, GetFriendsListPath)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIOptions.btnAddFriend_Accepted : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Sub

Private Sub lstFriends_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Call CommonListChange(lstFriends, nX, nY)
End Sub

Private Sub btnAddDetect_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    If g_ui.Loot.AddItemToList(lstDetect, txtAddDetect.Text, True, False) Then
        Call g_ui.Loot.SaveItemsList(lstDetect, GetDetectListPath)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIOptions.btnAddDetect_Accepted : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Sub

Private Sub lstDetect_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Call CommonListChange(lstDetect, nX, nY)
End Sub

Private Sub btnSetDetect_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    
    If g_Hooks.CurrentSelection <> 0 Then
        If g_Objects.Exists(g_Hooks.CurrentSelection) Then
            txtAddDetect.Text = g_Objects.FindObject(g_Hooks.CurrentSelection).Name
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "btnSetDetect_Accepted - " & Err.Description
    Resume Fin
End Sub


'-- HUD changes --
Private Sub chkEnableHUD_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        Call g_HUD.StartHUDs
    Else
        PrintMessage "Disabling All HUDs"
        Call g_HUD.stopAllHUDs
    End If
End Sub

Private Sub chkHUDMacro_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        PrintMessage "Enabling Macro HUD"
        Call g_HUD.startMacroHUD
    Else
        PrintMessage "Disabling Macro HUD"
        Call g_HUD.stopMacroHUD
    End If
End Sub
Private Sub chkHUDStats_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        PrintMessage "Enabling Stats HUD"
        Call g_HUD.startStatsHUD
    Else
        PrintMessage "Disabling Stats HUD"
        Call g_HUD.stopStatsHUD
    End If
End Sub
Private Sub chkHUDInventory_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        PrintMessage "Enabling Inventory HUD"
        Call g_HUD.startInvHUD
    Else
        PrintMessage "Disabling Inventory HUD"
        Call g_HUD.stopInvHUD
    End If
End Sub
Private Sub chkHUDDebug_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        PrintMessage "Enabling Debug HUD"
        Call g_HUD.startDebugHUD
    Else
        PrintMessage "Disabling Debug HUD"
        Call g_HUD.stopDebugHUD
    End If
End Sub
Private Sub chkHUDDOT_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        PrintMessage "Enabling DoT HUD"
        Call g_HUD.startDOTHUD
    Else
        PrintMessage "Disabling DoT HUD"
        Call g_HUD.stopDOTHUD
    End If
End Sub

'chkLockHuds
Private Sub chkLockHuds_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If bChecked Then
        PrintMessage "Saving HUD position"
        Call g_Settings.SavePluginConfiguration
        
        If Valid(g_HUD.getMacroHUD) Then
            g_ui.Options.macroHUDposX = g_HUD.getMacroHUD.Region.Left
            g_ui.Options.macroHUDposY = g_HUD.getMacroHUD.Region.Top
        End If
        If Valid(g_HUD.getStatsHUD) Then
            g_ui.Options.statsHUDposX = g_HUD.getStatsHUD.Region.Left
            g_ui.Options.statsHUDposY = g_HUD.getStatsHUD.Region.Top
        End If
        If Valid(g_HUD.getInvHUD) Then
            g_ui.Options.invHUDposX = g_HUD.getInvHUD.Region.Left
            g_ui.Options.invHUDposY = g_HUD.getInvHUD.Region.Top
        End If
        If Valid(g_HUD.getDebugHUD) Then
            g_ui.Options.debugHUDposX = g_HUD.getDebugHUD.Region.Left
            g_ui.Options.debugHUDposY = g_HUD.getDebugHUD.Region.Top
        End If
        If Valid(g_HUD.getDOTHUD) Then
            g_ui.Options.dotHUDposX = g_HUD.getDOTHUD.Region.Left
            g_ui.Options.dotHUDposY = g_HUD.getDOTHUD.Region.Top
        End If
    End If
End Sub

