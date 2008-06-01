VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIBuffs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================
'User Interface for tab : Buffs
'===========================================
Option Explicit

Private InterfaceName As String


'Settings tab
Public WithEvents chkRebuffOnStart As DecalControls.CheckBox
Attribute chkRebuffOnStart.VB_VarHelpID = -1
Public WithEvents chkEnableBuffing As DecalControls.CheckBox
Attribute chkEnableBuffing.VB_VarHelpID = -1
Public WithEvents chkContinuousBuffing As DecalControls.CheckBox
Attribute chkContinuousBuffing.VB_VarHelpID = -1
Public WithEvents chMaxBuffSpellLevel As DecalControls.Choice
Attribute chMaxBuffSpellLevel.VB_VarHelpID = -1

Public WithEvents txtRebuffInterval As DecalControls.Edit
Attribute txtRebuffInterval.VB_VarHelpID = -1
Public WithEvents txtNumContinuousBuffs As DecalControls.Edit
Attribute txtNumContinuousBuffs.VB_VarHelpID = -1

Public WithEvents chkEnableBuffBuddy As DecalControls.CheckBox
Attribute chkEnableBuffBuddy.VB_VarHelpID = -1
Public WithEvents btnSetBuffBuddy As DecalControls.PushButton
Attribute btnSetBuffBuddy.VB_VarHelpID = -1
Public WithEvents btnForceBuffBuddy As DecalControls.PushButton
Attribute btnForceBuffBuddy.VB_VarHelpID = -1
Public txtBuffBuddy As DecalControls.StaticText

'Creature/Life spells
Public WithEvents lstCreature As DecalControls.list
Attribute lstCreature.VB_VarHelpID = -1
Public WithEvents lstLifePros As DecalControls.list
Attribute lstLifePros.VB_VarHelpID = -1
Public chkEnableCreatureBuffs As DecalControls.CheckBox
Public chkEnableLifeBuffs As DecalControls.CheckBox
'Buff Buddy
Public WithEvents lstOtherCreature As DecalControls.list
Attribute lstOtherCreature.VB_VarHelpID = -1
Public WithEvents lstOtherLifePros As DecalControls.list
Attribute lstOtherLifePros.VB_VarHelpID = -1
Public chkEnableOtherCreatureBuffs As DecalControls.CheckBox
Public chkEnableOtherLifeBuffs As DecalControls.CheckBox

'Item
Public chkItmArmorHead As DecalControls.CheckBox
Public chkItmArmorHands As DecalControls.CheckBox
Public chkItmArmorTop As DecalControls.CheckBox
Public chkItmArmorBottom As DecalControls.CheckBox
Public chkItmArmorFeet As DecalControls.CheckBox
Public chkItmArmorTopUndie As DecalControls.CheckBox
Public chkItmArmorBottomUndie As DecalControls.CheckBox
Public chkItmArmorShield As DecalControls.CheckBox

Public chkItmBaneImpen As DecalControls.CheckBox
Public chkItmBaneSlash As DecalControls.CheckBox
Public chkItmBanePierce As DecalControls.CheckBox
Public chkItmBaneBludg As DecalControls.CheckBox
Public chkItmBaneFire As DecalControls.CheckBox
Public chkItmBaneAcid As DecalControls.CheckBox
Public chkItmBaneFrost As DecalControls.CheckBox
Public chkItmBaneLightning As DecalControls.CheckBox

Public chkItmWeapBD As DecalControls.CheckBox
Public chkItmWeapHS As DecalControls.CheckBox
Public chkItmWeapSK As DecalControls.CheckBox
Public chkItmWeapDE As DecalControls.CheckBox

Public chkItmWandHL As DecalControls.CheckBox
Public chkItmWandDef As DecalControls.CheckBox
Public chkItmWandSD As DecalControls.CheckBox

Public chkItmSelectSelf As DecalControls.CheckBox

'Buff buddy items
Public chkOtherItmArmorShield As DecalControls.CheckBox
Public chkOtherItmWeapBD As DecalControls.CheckBox
Public chkOtherItmWeapHS As DecalControls.CheckBox
Public chkOtherItmWeapSK As DecalControls.CheckBox
Public chkOtherItmWeapDE As DecalControls.CheckBox

Public chkOtherItmWandHL As DecalControls.CheckBox
Public chkOtherItmWandDef As DecalControls.CheckBox
Public chkOtherItmWandSD As DecalControls.CheckBox


'Extra
Public stItemName As DecalControls.StaticText
Public WithEvents btnSetItemGUID As DecalControls.PushButton
Attribute btnSetItemGUID.VB_VarHelpID = -1
Public WithEvents btnAddItem As DecalControls.PushButton
Attribute btnAddItem.VB_VarHelpID = -1
Public WithEvents lstItems As DecalControls.list
Attribute lstItems.VB_VarHelpID = -1
Public WithEvents lstItemSpells As DecalControls.list
Attribute lstItemSpells.VB_VarHelpID = -1
Public WithEvents chItemSpell As DecalControls.Choice
Attribute chItemSpell.VB_VarHelpID = -1
Public WithEvents btnAddSpellToItem As DecalControls.PushButton
Attribute btnAddSpellToItem.VB_VarHelpID = -1

Private objItemToAdd As acObject

'Controls declaration
Public Function Init(Optional ProfileName As String = "Default") As Boolean

    On Error GoTo Error_Handler

    'Initialize
    Init = False
    
    'Set interface module name
    InterfaceName = "UIBuffs"
    
    MyDebug InterfaceName & ".Init() -- Begin"
    
    'Decal Controls initialisation
    Set chkRebuffOnStart = g_MainView.Control("chkRebuffOnStart")
    Set chkEnableBuffing = g_MainView.Control("chkEnableBuffing")
    Set chkContinuousBuffing = g_MainView.Control("chkContinuousBuffing")
    
    'Set stItemName = g_MainView.Control("stItemName")
    'Set btnSetItemGUID = g_MainView.Control("btnSetItemGUID")
    'Set btnAddItem = g_MainView.Control("btnAddItem")
    'Set lstItemSpells = g_MainView.Control("lstItemSpells")
    'Set lstItems = g_MainView.Control("lstItems")
    'Set chItemSpell = g_MainView.Control("chItemSpell")
    'Set btnAddSpellToItem = g_MainView.Control("btnAddSpellToItem")
    
    Set chkEnableBuffBuddy = g_MainView.Control("chkEnableBuffBuddy")
    Set btnSetBuffBuddy = g_MainView.Control("btnSetBuffBuddy")
    Set btnForceBuffBuddy = g_MainView.Control("btnForceBuffBuddy")
    Set txtBuffBuddy = g_MainView.Control("txtBuffBuddy")
    Set lstLifePros = g_MainView.Control("lstLifePros")
    Set lstCreature = g_MainView.Control("lstCreature")
    Set lstOtherCreature = g_MainView.Control("lstOtherCreature")
    Set lstOtherLifePros = g_MainView.Control("lstOtherLifePros")
    
    Set txtRebuffInterval = g_MainView.Control("txtRebuffInterval")
    Set txtNumContinuousBuffs = g_MainView.Control("txtNumContinuousBuffs")
    
    Set chkItmArmorHead = g_MainView.Control("chkItmArmorHead")
    Set chkItmArmorHands = g_MainView.Control("chkItmArmorHands")
    Set chkItmArmorTop = g_MainView.Control("chkItmArmorTop")
    Set chkItmArmorBottom = g_MainView.Control("chkItmArmorBottom")
    Set chkItmArmorFeet = g_MainView.Control("chkItmArmorFeet")
    Set chkItmArmorFeet = g_MainView.Control("chkItmArmorFeet")
    Set chkItmArmorTopUndie = g_MainView.Control("chkItmArmorTopUndie")
    Set chkItmArmorBottomUndie = g_MainView.Control("chkItmArmorBottomUndie")
    Set chkItmArmorShield = g_MainView.Control("chkItmArmorShield")
    Set chkItmBaneImpen = g_MainView.Control("chkItmBaneImpen")
    Set chkItmBaneSlash = g_MainView.Control("chkItmBaneSlash")
    Set chkItmBanePierce = g_MainView.Control("chkItmBanePierce")
    Set chkItmBaneBludg = g_MainView.Control("chkItmBaneBludg")
    Set chkItmBaneFire = g_MainView.Control("chkItmBaneFire")
    Set chkItmBaneAcid = g_MainView.Control("chkItmBaneAcid")
    Set chkItmBaneFrost = g_MainView.Control("chkItmBaneFrost")
    Set chkItmBaneLightning = g_MainView.Control("chkItmBaneLightning")
    Set chkItmWeapBD = g_MainView.Control("chkItmWeapBD")
    Set chkItmWeapHS = g_MainView.Control("chkItmWeapHS")
    Set chkItmWeapSK = g_MainView.Control("chkItmWeapSK")
    Set chkItmWeapDE = g_MainView.Control("chkItmWeapDE")
    Set chkItmWandHL = g_MainView.Control("chkItmWandHL")
    Set chkItmWandDef = g_MainView.Control("chkItmWandDef")
    Set chkItmWandSD = g_MainView.Control("chkItmWandSD")
    Set chkItmSelectSelf = g_MainView.Control("chkItmSelectSelf")
    
    Set chkOtherItmArmorShield = g_MainView.Control("chkOtherItmArmorShield")
    Set chkOtherItmWeapBD = g_MainView.Control("chkOtherItmWeapBD")
    Set chkOtherItmWeapHS = g_MainView.Control("chkOtherItmWeapHS")
    Set chkOtherItmWeapSK = g_MainView.Control("chkOtherItmWeapSK")
    Set chkOtherItmWeapDE = g_MainView.Control("chkOtherItmWeapDE")
    Set chkOtherItmWandHL = g_MainView.Control("chkOtherItmWandHL")
    Set chkOtherItmWandDef = g_MainView.Control("chkOtherItmWandDef")
    Set chkOtherItmWandSD = g_MainView.Control("chkOtherItmWandSD")
    
    Set chMaxBuffSpellLevel = g_MainView.Control("chMaxBuffSpellLevel")
    
    Set chkEnableCreatureBuffs = g_MainView.Control("chkEnableCreatureBuffs")
    Set chkEnableLifeBuffs = g_MainView.Control("chkEnableLifeBuffs")
    Set chkEnableOtherCreatureBuffs = g_MainView.Control("chkEnableOtherCreatureBuffs")
    Set chkEnableOtherLifeBuffs = g_MainView.Control("chkEnableOtherLifeBuffs")

    
    'Controls default settings
    'chItemSpell.Selected = 0
    
    Init = True
    MyDebug InterfaceName & ".Init() -- End"
Fin:
    Exit Function
    
Error_Handler:
    Init = False
    PrintErrorMessage "(interface:" & InterfaceName & ") : Error #" & CStr(Err.Number) & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin

End Function

Public Function LoadControlsValue(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler

    MyDebug "[" & InterfaceName & "] Loading controls value"
    
     '***************************************************************************
    'chItemSpell.Selected = 0
    chkRebuffOnStart.Checked = g_Settings.GetValue("chkRebuffOnStart", True)
    chkEnableBuffing.Checked = g_Settings.GetValue("chkEnableBuffing", True)
    chkContinuousBuffing.Checked = g_Settings.GetValue("chkContinuousBuffing", False)
    txtRebuffInterval.Text = g_Settings.GetValue("txtRebuffInterval", 60)
    chMaxBuffSpellLevel.Selected = g_Settings.GetValue("chMaxBuffSpellLevel", 6) 'default to level 7
    txtNumContinuousBuffs.Text = g_Settings.GetValue("txtNumContinuousBuffs", 1)
    chkEnableBuffBuddy.Checked = g_Settings.GetValue("chkEnableBuffBuddy", False)
    
    With g_Settings.Profile.BuffCfg
        chkItmArmorHead.Checked = .GetValue("chkItmArmorHead", True)
        chkItmArmorHands.Checked = .GetValue("chkItmArmorHands", True)
        chkItmArmorTop.Checked = .GetValue("chkItmArmorTop", True)
        chkItmArmorBottom.Checked = .GetValue("chkItmArmorBottom", True)
        chkItmArmorFeet.Checked = .GetValue("chkItmArmorFeet", True)
        chkItmArmorTopUndie.Checked = .GetValue("chkItmArmorTopUndie", True)
        chkItmArmorBottomUndie.Checked = .GetValue("chkItmArmorBottomUndie", True)
        chkItmArmorShield.Checked = .GetValue("chkItmArmorShield", True)
        chkItmBaneImpen.Checked = .GetValue("chkItmBaneImpen", True)
        chkItmBaneSlash.Checked = .GetValue("chkItmBaneSlash", True)
        chkItmBanePierce.Checked = .GetValue("chkItmBanePierce", True)
        chkItmBaneBludg.Checked = .GetValue("chkItmBaneBludg", True)
        chkItmBaneFire.Checked = .GetValue("chkItmBaneFire", True)
        chkItmBaneAcid.Checked = .GetValue("chkItmBaneAcid", True)
        chkItmBaneFrost.Checked = .GetValue("chkItmBaneFrost", True)
        chkItmBaneLightning.Checked = .GetValue("chkItmBaneLightning", True)
        chkItmWeapBD.Checked = .GetValue("chkItmWeapBD", True)
        chkItmWeapHS.Checked = .GetValue("chkItmWeapHS", True)
        chkItmWeapSK.Checked = .GetValue("chkItmWeapSK", True)
        chkItmWeapDE.Checked = .GetValue("chkItmWeapDE", True)
        chkItmWandHL.Checked = .GetValue("chkItmWandHL", True)
        chkItmWandDef.Checked = .GetValue("chkItmWandDef", True)
        chkItmWandSD.Checked = .GetValue("chkItmWandSD", True)
        chkItmSelectSelf.Checked = .GetValue("chkItmSelectSelf", True)
        
        chkOtherItmArmorShield.Checked = .GetValue("chkOtherItmArmorShield", True)
        chkOtherItmWeapBD.Checked = .GetValue("chkOtherItmWeapBD", True)
        chkOtherItmWeapHS.Checked = .GetValue("chkOtherItmWeapHS", True)
        chkOtherItmWeapSK.Checked = .GetValue("chkOtherItmWeapSK", True)
        chkOtherItmWeapDE.Checked = .GetValue("chkOtherItmWeapDE", True)
        chkOtherItmWandHL.Checked = .GetValue("chkOtherItmWandHL", True)
        chkOtherItmWandDef.Checked = .GetValue("chkOtherItmWandDef", True)
        chkOtherItmWandSD.Checked = .GetValue("chkOtherItmWandSD", True)

    End With

    chkEnableCreatureBuffs.Checked = g_Settings.GetValue("chkEnableCreatureBuffs", True)
    chkEnableLifeBuffs.Checked = g_Settings.GetValue("chkEnableLifeBuffs", True)
    
    chkEnableOtherCreatureBuffs.Checked = g_Settings.GetValue("chkEnableOtherCreatureBuffs", True)
    chkEnableOtherLifeBuffs.Checked = g_Settings.GetValue("chkEnableOtherLifeBuffs", True)
  
    Call InitBuffsListControls
    '***************************************************************************
    
    LoadControlsValue = True
    
Fin:
    Exit Function
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage "clsUIBuffs.LoadControlsValue) - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    Dim i As Integer
    
    MyDebug "[" & InterfaceName & "] Saving controls settings"
    
    With g_Settings.Profile.Config
        .SaveCheckbox chkRebuffOnStart, "chkRebuffOnStart"
        .SaveCheckbox chkEnableBuffing, "chkEnableBuffing"
        .SaveCheckbox chkContinuousBuffing, "chkContinuousBuffing"
        .SaveCheckbox chkEnableCreatureBuffs, "chkEnableCreatureBuffs"
        .SaveCheckbox chkEnableLifeBuffs, "chkEnableLifeBuffs"
        .SaveCheckbox chkEnableOtherCreatureBuffs, "chkEnableOtherCreatureBuffs"
        .SaveCheckbox chkEnableOtherLifeBuffs, "chkEnableOtherLifeBuffs"
        .SaveTextbox txtRebuffInterval, "txtRebuffInterval"
        .SaveTextbox txtNumContinuousBuffs, "txtNumContinuousBuffs"
        .SaveChoice chMaxBuffSpellLevel, "chMaxBuffSpellLevel"
        .SaveCheckbox chkEnableBuffBuddy, "chkEnableBuffBuddy"
    End With
    
    With g_Settings.Profile.BuffCfg
        'Item spells/parts
        .SaveCheckbox chkItmArmorHead, "chkItmArmorHead"
        .SaveCheckbox chkItmArmorHands, "chkItmArmorHands"
        .SaveCheckbox chkItmArmorTop, "chkItmArmorTop"
        .SaveCheckbox chkItmArmorBottom, "chkItmArmorBottom"
        .SaveCheckbox chkItmArmorFeet, "chkItmArmorFeet"
        .SaveCheckbox chkItmArmorTopUndie, "chkItmArmorTopUndie"
        .SaveCheckbox chkItmArmorBottomUndie, "chkItmArmorBottomUndie"
        .SaveCheckbox chkItmArmorShield, "chkItmArmorShield"
        .SaveCheckbox chkItmBaneImpen, "chkItmBaneImpen"
        .SaveCheckbox chkItmBaneSlash, "chkItmBaneSlash"
        .SaveCheckbox chkItmBanePierce, "chkItmBanePierce"
        .SaveCheckbox chkItmBaneBludg, "chkItmBaneBludg"
        .SaveCheckbox chkItmBaneFire, "chkItmBaneFire"
        .SaveCheckbox chkItmBaneAcid, "chkItmBaneAcid"
        .SaveCheckbox chkItmBaneFrost, "chkItmBaneFrost"
        .SaveCheckbox chkItmBaneLightning, "chkItmBaneLightning"
        .SaveCheckbox chkItmWeapBD, "chkItmWeapBD"
        .SaveCheckbox chkItmWeapHS, "chkItmWeapHS"
        .SaveCheckbox chkItmWeapSK, "chkItmWeapSK"
        .SaveCheckbox chkItmWeapDE, "chkItmWeapDE"
        .SaveCheckbox chkItmWandHL, "chkItmWandHL"
        .SaveCheckbox chkItmWandDef, "chkItmWandDef"
        .SaveCheckbox chkItmWandSD, "chkItmWandSD"
        .SaveCheckbox chkItmSelectSelf, "chkItmSelectSelf"
        
        .SaveCheckbox chkOtherItmArmorShield, "chkOtherItmArmorShield"
        .SaveCheckbox chkOtherItmWeapBD, "chkOtherItmWeapBD"
        .SaveCheckbox chkOtherItmWeapHS, "chkOtherItmWeapHS"
        .SaveCheckbox chkOtherItmWeapSK, "chkOtherItmWeapSK"
        .SaveCheckbox chkOtherItmWeapDE, "chkOtherItmWeapDE"
        .SaveCheckbox chkOtherItmWandHL, "chkOtherItmWandHL"
        .SaveCheckbox chkOtherItmWandDef, "chkOtherItmWandDef"
        .SaveCheckbox chkOtherItmWandSD, "chkOtherItmWandSD"
        
        'Save Creature spells
        For i = 0 To lstCreature.Count - 1
            Call .SaveValue(lstCreature.Data(1, i, 0), BoolToInteger(lstCreature.Data(2, i, 0)))
        Next i
    
        'Save Life spells
        For i = 0 To lstLifePros.Count - 1
            Call .SaveValue(lstLifePros.Data(1, i, 0), BoolToInteger(lstLifePros.Data(2, i, 0)))
        Next i
        
        'Save Buff Buddy Creature Spells
        For i = 0 To lstOtherCreature.Count - 1
            Call .SaveValue(lstOtherCreature.Data(1, i, 0), BoolToInteger(lstOtherCreature.Data(2, i, 0)))
        Next i
        
        'Save Buff Buddy Life spells
        For i = 0 To lstOtherLifePros.Count - 1
            Call .SaveValue(lstOtherLifePros.Data(1, i, 0), BoolToInteger(lstOtherLifePros.Data(2, i, 0)))
        Next i
        
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
    
    'Unload initialization
    Unload = False

    Set objItemToAdd = Nothing
    Set chkRebuffOnStart = Nothing
    Set chkEnableBuffing = Nothing
    Set chkEnableBuffBuddy = Nothing
    Set chkContinuousBuffing = Nothing
    
    Set stItemName = Nothing
    Set btnSetItemGUID = Nothing
    Set btnAddItem = Nothing
    Set lstItemSpells = Nothing
    Set lstItems = Nothing
    Set chItemSpell = Nothing
    Set btnAddSpellToItem = Nothing
    
    Set btnSetBuffBuddy = Nothing
    Set btnForceBuffBuddy = Nothing
    Set txtBuffBuddy = Nothing
    Set lstLifePros = Nothing
    Set lstCreature = Nothing
    Set lstOtherCreature = Nothing
    Set lstOtherLifePros = Nothing
    Set txtRebuffInterval = Nothing
    Set chkItmArmorHead = Nothing
    Set chkItmArmorHands = Nothing
    Set chkItmArmorTop = Nothing
    Set chkItmArmorBottom = Nothing
    Set chkItmArmorFeet = Nothing
    Set chkItmArmorTopUndie = Nothing
    Set chkItmArmorBottomUndie = Nothing
    Set chkItmArmorShield = Nothing
    Set chkItmBaneImpen = Nothing
    Set chkItmBaneSlash = Nothing
    Set chkItmBanePierce = Nothing
    Set chkItmBaneBludg = Nothing
    Set chkItmBaneFire = Nothing
    Set chkItmBaneAcid = Nothing
    Set chkItmBaneFrost = Nothing
    Set chkItmBaneLightning = Nothing
    Set chkItmWeapBD = Nothing
    Set chkItmWeapHS = Nothing
    Set chkItmWeapSK = Nothing
    Set chkItmWeapDE = Nothing
    Set chkItmWandHL = Nothing
    Set chkItmWandDef = Nothing
    Set chkItmWandSD = Nothing
    Set chkItmSelectSelf = Nothing
    Set chMaxBuffSpellLevel = Nothing
    Set txtNumContinuousBuffs = Nothing
    
    Set chkOtherItmArmorShield = Nothing
    Set chkOtherItmWeapBD = Nothing
    Set chkOtherItmWeapHS = Nothing
    Set chkOtherItmWeapSK = Nothing
    Set chkOtherItmWeapDE = Nothing
    Set chkOtherItmWandHL = Nothing
    Set chkOtherItmWandDef = Nothing
    Set chkOtherItmWandSD = Nothing
    
    Set chkEnableLifeBuffs = Nothing
    Set chkEnableCreatureBuffs = Nothing
    Set chkEnableOtherCreatureBuffs = Nothing
    Set chkEnableOtherLifeBuffs = Nothing
    
    'Unload complete
    Unload = True
    MyDebug InterfaceName & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage "(interface:" & InterfaceName & ") - " & Err.Description
    Resume Fin
End Function


Private Sub Class_Terminate()
    Call Unload
End Sub


'==========================================================================================
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'==========================================================================================

Private Sub btnSetBuffBuddy_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    If g_Hooks.CurrentSelection <> 0 Then
        Dim obj As acObject
        Set obj = g_Objects.FindObject(g_Hooks.CurrentSelection)
        If Valid(obj) And IsPlayer(obj) Then
            Set g_buffBuddy = obj
            txtBuffBuddy.Text = g_buffBuddy.Name
            
            'Find and set wielded items for our buff buddy
            Dim objItem As acObject
            For Each objItem In g_Objects.Items.World
                If (objItem.Wielder = g_buffBuddy.Guid) And (objItem.Container = 0) Then
                    If objItem.itemType = ITEM_ARMOR Then
                        'It's a shield
                        MyDebug "Found Buff Buddy's shield: " & objItem.Name
                        Call g_BuddyBuffer.setHeldShield(objItem)
                    Else
                        'It's a weapon or wand
                        MyDebug "Found Buff Buddy's wand/weap: " & objItem.Name
                        Call g_BuddyBuffer.setHeldItem(objItem)
                    End If
                End If
            Next objItem
            
            Dim iBuffMode As Integer
                    
            If g_ui.Buffs.chkContinuousBuffing.Checked Then
                iBuffMode = REBUFF_CONTINUOUS
            Else
                iBuffMode = REBUFF_FULL
            End If
            
            Call g_BuddyBuffer.StartService(iBuffMode, g_Data.RebuffInterval, True, g_ui.Buffs.chkRebuffOnStart.Checked, g_Data.NumContinuousBuffs)

        Else
            txtBuffBuddy.Text = "NONE"
            PrintMessage "Error, please select a different object"
        End If
    Else
        txtBuffBuddy.Text = "NONE"
        PrintMessage "Please select Player to set as Buff Buddy"
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIBuffs.btnSetBuffBuddy - " & Err.Description
    Resume Fin
End Sub

Private Sub btnForceBuffBuddy_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    Call g_Macro.ForceBuddyBuff

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIBuffs.btnForceBuffBuddy - " & Err.Description
    Resume Fin
End Sub

'Private Sub btnApplyBuffsChanges_Accepted(ByVal nID As Long)
'
'
'    Dim NewRebuffInterval As Long
'    Dim NewNumContinuousBuffs As Integer
'    Dim Msg As String
'
'    NewRebuffInterval = Val(txtRebuffInterval.Text)
'    NewNumContinuousBuffs = Val(txtNumContinuousBuffs.Text)
'
'    If NewRebuffInterval > 0 Then
'        Macro.RebuffInterval = 60 * NewRebuffInterval * TIMESCALE
'
'        'Needed for Continuous Buff mode (ItemBuffNum)
'        BuildItemBuffsList
'
'        'Update Buff interval for continuous buffing mode
'        SetContinuousBuffInterval
'
'        'MyDebug "ResetMacro: Macro.ContinuousBuffingInerval = " & Macro.ContinuousBuffingInterval
'        If ContinuousBuffingMode = True Then
'            PrintMessage "Continuous Buffing Mode: Macro will cast a new buff spell every " & Macro.ContinuousBuffingInterval & " seconds"
'        Else
'            PrintMessage "Macro will now rebuff every " & NewRebuffInterval & " minute(s)"
'        End If
'
'    End If
'
'    If (NewNumContinuousBuffs > 0) And (NewNumContinuousBuffs <= (GetTotalEnabledBuffs + ItemBuffList.NumBuffs)) Then 'fixme: make sure this value is not higher than our current total num of buffs
'        Macro.ContinuousBuffNum = NewNumContinuousBuffs
'        'Update Buff interval for continuous buffing mode
'        SetContinuousBuffInterval
'
'        PrintMessage "Macro will now cast " & NewNumContinuousBuffs & " spells consecutively while in Continuous Buffing mode."
'
'    Else
'        PrintMessage "Invalid Num Consecutive buffs."
'    End If
'
'    SavePluginConfiguration
'
'End Sub
'
'Private Sub btnSaveNewBuffProfile_Accepted(ByVal nID As Long)
'Dim NewProfileName As String
'
'    NewProfileName = Trim(txtNewBuffProfile.Text)
'    If NewProfileName <> "" Then
'        If ProfileExists(NewProfileName) Then
'            PrintMessage "This buff profile already exists. Please choose another buff profile name."
'        Else
'            Call CreateNewBuffProfile(NewProfileName, g_ui.ProfileName)
'            Call SavePluginConfiguration
'            'Call g_ui.LoadBuffProfile(g_ui.ProfileName, NewProfileName)
'            Call g_ui.Main.UpdateBuffProfileList
'            Call ChoiceListSelect(NewProfileName, g_ui.Main.chBuffProfile)
'            PrintMessage "New Buff Profile saved as : " & NewProfileName
'        End If
'
'    Else
'        PrintMessage "Please enter a valid profile name"
'    End If
'End Sub
'
'
'
'
'Private Sub chMaxBuffSpellLevel_Change(ByVal nID As Long, ByVal nIndex As Long)
'    If nIndex < 0 Or chMaxBuffSpellLevel.Selected < 0 Then
'        Exit Sub
'    End If
'
'    Call UpdateMaxBuffSpellLevel
'    PrintMessage "Buff Spells Level set to " & Macro.Spells.MaxBuffLevel
'
'End Sub

'
'Private Sub btnSetItemGUID_Accepted(ByVal nID As Long)
'
'    If Not Macro.Items.Exists(PluginSite.CurrentSelection) Then
'        PrintMessage "Please select a valid item (armor/weapon) you want to add spells on, and click Set, then Add to add it to the items list."
'    Else
'        Set objItemToAdd = Macro.Items(PluginSite.CurrentSelection)
'        stItemGUID.Text = objItemToAdd.GUID
'        stItemName.Text = objItemToAdd.Name
'    End If
'
'End Sub




Private Sub AddBuffToList(ByRef lstControl As DecalControls.list, ByVal sSpellFamily As String, Optional ByVal Icon As Long = 0, Optional ByVal bChecked As Boolean = True)
On Error GoTo Error_Handler
    
    Dim i As Integer

    i = lstControl.AddRow
    lstControl.Data(0, i, 1) = Icon
    lstControl.Data(1, i, 0) = sSpellFamily
    lstControl.Data(2, i, 0) = g_Settings.Profile.BuffCfg.GetValue(sSpellFamily, bChecked)   'checked or not?
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage InterfaceName & ".AddBuffToList - " & Err.Description
    Resume Fin
End Sub

Public Sub InitBuffsListControls()
On Error GoTo ErrorHandler

    MyDebug "Filling Buff Spells lists controls (InitBuffsListControls)"

    Dim objFam As clsSpellFamily
    
    Call lstCreature.Clear
    Call lstLifePros.Clear
    For Each objFam In g_Spells.sBuffs.Families
        If (objFam.mSchool = SCHOOL_LIFE) And (objFam.mType = SPELLTYPE_LIFEPRO) Then
            AddBuffToList lstLifePros, objFam.mFamily, objFam.mIcon
        Else
            'creature & non life-prot life spells
            AddBuffToList lstCreature, objFam.mFamily, objFam.mIcon
        End If
    Next objFam
    
    Call lstOtherCreature.Clear
    Call lstOtherLifePros.Clear
    For Each objFam In g_Spells.oBuffs.Families
        If (objFam.mSchool = SCHOOL_LIFE) And (objFam.mType = SPELLTYPE_LIFEPRO) Then
            MyDebug "clsUIBuffs.InitBuffsListControls: lsOtherLifePros: " & objFam.mFamily
            AddBuffToList lstOtherLifePros, objFam.mFamily, objFam.mIcon
        Else
            'creature & non life-prot life spells
            AddBuffToList lstOtherCreature, objFam.mFamily, objFam.mIcon
        End If
    Next objFam
    
    'Call chItemSpell.Clear
    'For Each objFam In g_Spells.Items.Families
    '    If objFam.mType <> SPELLTYPE_TRANSPORT Then  'only add regular spells to the list, no recalls/ties etc
    '        chItemSpell.AddChoice objFam.mFamily
    '    End If
    'Next objFam
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "InitBuffsListControls - " & Err.Description
    Resume Fin
End Sub

