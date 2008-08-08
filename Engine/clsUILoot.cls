VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUILoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================
' Macro > Loot panel
Private Const INTERFACE_NAME = "UILoot"
'===========================================

Option Explicit

Private Enum eArmorFilterCols
    COL_ARMOR_ENABLE = 0
    COL_ARMOR_TYPE
    COL_ARMOR_COVERAGE
    COL_ARMOR_MIN_AL
    COL_ARMOR_MAX_WORK
    COL_ARMOR_MAX_VALUE
End Enum

Private Enum eWeaponFilterCols
    COL_WEAPON_ENABLE = 0
    COL_WEAPON_TYPE
    COL_WEAPON_ELEMENT
    COL_WEAPON_DMG
    COL_WEAPON_ATT
    COL_WEAPON_DEF
    COL_WEAPON_MAX_WORK
    COL_WEAPON_WIELD_REQ
    COL_WEAPON_MAX_VALUE
End Enum

Private Enum eWandFilterCols
    COL_WAND_ENABLE = 0
    COL_WAND_TYPE
    COL_WAND_MANAC
    COL_WAND_MELEED
    COL_WAND_MAGICD
    COL_WAND_PVM
    COL_WAND_WIELD_REQ
    COL_WAND_MAX_WORK
    COL_WAND_MAX_VALUE
End Enum

Private Enum eSalvageFilterCols
    COL_SALVAGE_ENABLE = 0
    COL_SALVAGE_NAME
    COL_SALVAGE_MIN_WORK
    COL_SALVAGE_COMBINE
End Enum

Private Enum eWeaponKindTabs
    TAB_MELEE_WEAPONS
    TAB_MISSILE_WEAPONS
End Enum

Private Enum ePickupListColumns
    COL_ENABLED
    COL_ITEM_NAME
    COL_EXACT_MATCH
    COL_NUM_PICKUP
End Enum

Private Enum eExceptionListColumns
    EX_COL_ICON = 0
    EX_COL_NAME
    EX_COL_GUID
End Enum

Private Const m_COLOR_DEL_COL = 15383387

'Settings tab -----------------------------------------------------

Public WithEvents chkEnableLooting As DecalControls.CheckBox
Attribute chkEnableLooting.VB_VarHelpID = -1
Public WithEvents chkEnableSalvage As DecalControls.CheckBox
Attribute chkEnableSalvage.VB_VarHelpID = -1

Public chkSalvageFrequency As DecalControls.CheckBox
Public txtSalvageFrequency As DecalControls.Edit
Public chkEnableMultipleSalvage As DecalControls.CheckBox

Public WithEvents txtPickupRange As DecalControls.Edit
Attribute txtPickupRange.VB_VarHelpID = -1
Public WithEvents chkLootFellowshipKills As DecalControls.CheckBox
Attribute chkLootFellowshipKills.VB_VarHelpID = -1
Public WithEvents chkLootGround As DecalControls.CheckBox
Attribute chkLootGround.VB_VarHelpID = -1
Public WithEvents chkLootForcePeaceMode As DecalControls.CheckBox
Attribute chkLootForcePeaceMode.VB_VarHelpID = -1

Public WithEvents chkSendRares As DecalControls.CheckBox
Attribute chkSendRares.VB_VarHelpID = -1
Public WithEvents chkSendName As DecalControls.CheckBox
Attribute chkSendName.VB_VarHelpID = -1

Public WithEvents chkBoostLootPriority As DecalControls.CheckBox
Attribute chkBoostLootPriority.VB_VarHelpID = -1

'Global filters
Public chkLootAny As DecalControls.CheckBox

Public chkPickupValuable As DecalControls.CheckBox
Public txtPickupMinValue As DecalControls.Edit
    
Public chkBurdenRatio As DecalControls.CheckBox
Public txtBurdenRatio As DecalControls.Edit

Public chkPickupMinors As DecalControls.CheckBox
    Public chkMinorWard As DecalControls.CheckBox
    Public chkMinorAttribute As DecalControls.CheckBox
    Public chkMinorMastery As DecalControls.CheckBox

Public chkPickupMajors As DecalControls.CheckBox
    Public chkMajorIgnoreBane As DecalControls.CheckBox

Public chkUnknownScrolls As DecalControls.CheckBox

'Salvages Filter
Public chkLootSalvages As DecalControls.CheckBox
Public chSalvage As DecalControls.Choice
Public WithEvents cmdAddSalvage As DecalControls.PushButton
Attribute cmdAddSalvage.VB_VarHelpID = -1
Public WithEvents chSalvageMinWork As DecalControls.Choice
Attribute chSalvageMinWork.VB_VarHelpID = -1
Public WithEvents chSalvageSortWork As DecalControls.Choice
Attribute chSalvageSortWork.VB_VarHelpID = -1
Public WithEvents lstLootSalvages As DecalControls.list
Attribute lstLootSalvages.VB_VarHelpID = -1


'Armor Filters
Public chkLootArmors As DecalControls.CheckBox
Public WithEvents lstLootArmors As DecalControls.list
Attribute lstLootArmors.VB_VarHelpID = -1
Public WithEvents chArmorType As DecalControls.Choice
Attribute chArmorType.VB_VarHelpID = -1
Public WithEvents chCoverage As DecalControls.Choice
Attribute chCoverage.VB_VarHelpID = -1
Public WithEvents chMaxArmorWork As DecalControls.Choice
Attribute chMaxArmorWork.VB_VarHelpID = -1
Public txtArmorMinAL As DecalControls.Edit
Public txtArmorMaxValue As DecalControls.Edit
Public WithEvents cmdAddArmor As DecalControls.PushButton
Attribute cmdAddArmor.VB_VarHelpID = -1

'Weapons Filters
Public chkLootWeapons As DecalControls.CheckBox
Public WithEvents lstLootWeapons As DecalControls.list
Attribute lstLootWeapons.VB_VarHelpID = -1
Public WithEvents chWeaponMaxWork As DecalControls.Choice
Attribute chWeaponMaxWork.VB_VarHelpID = -1
Public txtWeaponMinDef As DecalControls.Edit
Public chMaxWieldReq As DecalControls.Choice
Public txtWeaponMaxValue As DecalControls.Edit
Public WithEvents nbWeaponKind As DecalControls.Notebook
Attribute nbWeaponKind.VB_VarHelpID = -1

'Melee Weapons tab
Public chMeleeWeapons As DecalControls.Choice
Public chWeaponElement As DecalControls.Choice
Public txtWeaponMinLowDmg As DecalControls.Edit
Public txtWeaponMinHighDmg As DecalControls.Edit
Public txtWeaponMinAtt As DecalControls.Edit
    
'Missile Weapons tab
Public WithEvents chMissileWeapons As DecalControls.Choice
Attribute chMissileWeapons.VB_VarHelpID = -1
Public chMissileWeaponElement As DecalControls.Choice
Public txtWeaponMinMod As DecalControls.Edit
Public chMinDamageBonus As DecalControls.Choice
Public WithEvents cmdAddWeapon As DecalControls.PushButton
Attribute cmdAddWeapon.VB_VarHelpID = -1

'Wand Filters
Public chkLootWands As DecalControls.CheckBox
Public WithEvents lstLootWands As DecalControls.list
Attribute lstLootWands.VB_VarHelpID = -1
Public chWandMinManaC As DecalControls.Choice
Public chWandMinMeleeDef As DecalControls.Choice
Public chWandMinMagicDef As DecalControls.Choice
Public chWandMinPvMBonus As DecalControls.Choice
Public chWandMaxWieldReq As DecalControls.Choice
Public chWandMaxWork As DecalControls.Choice
Public WithEvents cmdAddWand As DecalControls.PushButton
Attribute cmdAddWand.VB_VarHelpID = -1

'Items tab ---------------------------------------------------------
Public txtAddItemToPickup As DecalControls.Edit
Public txtNumToPickup As DecalControls.Edit
Public WithEvents btnAddItemToPickup As DecalControls.PushButton
Attribute btnAddItemToPickup.VB_VarHelpID = -1
Public WithEvents btnPickupSetSelection As DecalControls.PushButton
Attribute btnPickupSetSelection.VB_VarHelpID = -1
Public WithEvents lstPickup As DecalControls.list
Attribute lstPickup.VB_VarHelpID = -1

Public txtAddItemToIgnore As DecalControls.Edit
Public WithEvents btnAddItemToIgnore As DecalControls.PushButton
Attribute btnAddItemToIgnore.VB_VarHelpID = -1
Public WithEvents btnIgnoreSetSelection As DecalControls.PushButton
Attribute btnIgnoreSetSelection.VB_VarHelpID = -1
Public WithEvents lstIgnore As DecalControls.list
Attribute lstIgnore.VB_VarHelpID = -1

'Corpses tab ---------------------------------------------------------
Public WithEvents lstCorpse As DecalControls.list
Attribute lstCorpse.VB_VarHelpID = -1
Public WithEvents chkFilterCorpses As DecalControls.CheckBox
Attribute chkFilterCorpses.VB_VarHelpID = -1

Public txtAddCorpse As DecalControls.Edit
Public WithEvents btnAddCorpse As DecalControls.PushButton
Attribute btnAddCorpse.VB_VarHelpID = -1
Public WithEvents btnCorpseSetSelection As DecalControls.PushButton
Attribute btnCorpseSetSelection.VB_VarHelpID = -1

'Exceptions tab ---------------------------------------------------------
Public WithEvents lstExceptions As DecalControls.list
Attribute lstExceptions.VB_VarHelpID = -1
Public WithEvents cmdAddException As DecalControls.PushButton
Attribute cmdAddException.VB_VarHelpID = -1

'Mana Stones tab--------------------------------------------------------
Public txtHighMana As DecalControls.Edit

'-----------------------------------------------------------------

Private Const TAG_ITEM_NAME = "item"
Private Const TAG_ENABLED = "enable"
Private Const TAG_EXACT_MATCH = "exact"
Private Const TAG_NUM_PICKUP = "numpickup"

'Associate a list line index to a filter Id
Private m_dicRowToFilterId(0 To NUM_LOOT_FILTERS) As Dictionary

Private Sub Class_Initialize()
    Dim i As Integer
    For i = 0 To NUM_LOOT_FILTERS
        Set m_dicRowToFilterId(i) = New Dictionary
    Next i
End Sub

Private Sub Class_Terminate()
    Call Unload
    Dim i As Integer
    For i = 0 To NUM_LOOT_FILTERS
        Set m_dicRowToFilterId(i) = Nothing
    Next i
End Sub

Public Function Init() As Boolean
On Error GoTo Error_Handler

    'Initialize
    Init = False
    
    MyDebug INTERFACE_NAME & ".Init() -- Begin"
    
    'Create Controls
    Set chkEnableLooting = g_MainView.Control("chkEnableLooting")
    Set chkEnableSalvage = g_MainView.Control("chkEnableSalvage")

    Set txtPickupRange = g_MainView.Control("txtPickupRange")
    
    Set txtPickupMinValue = g_MainView.Control("txtPickupMinValue")
    
    Set txtHighMana = g_MainView.Control("txtHighMana")
    
    Set txtAddItemToPickup = g_MainView.Control("txtAddItemToPickup")
    Set btnAddItemToPickup = g_MainView.Control("btnAddItemToPickup")
    Set txtNumToPickup = g_MainView.Control("txtNumToPickup")
    Set btnPickupSetSelection = g_MainView.Control("btnPickupSetSelection")
    Set lstPickup = g_MainView.Control("lstPickup")
    
    Set txtAddItemToIgnore = g_MainView.Control("txtAddItemToIgnore")
    Set btnAddItemToIgnore = g_MainView.Control("btnAddItemToIgnore")
    Set btnIgnoreSetSelection = g_MainView.Control("btnIgnoreSetSelection")
    Set lstIgnore = g_MainView.Control("lstIgnore")

    Set chkLootAny = g_MainView.Control("chkLootAny")
    Set chkPickupValuable = g_MainView.Control("chkPickupValuable")
    Set chkBurdenRatio = g_MainView.Control("chkBurdenRatio")
    Set txtBurdenRatio = g_MainView.Control("txtBurdenRatio")
    Set chkPickupMinors = g_MainView.Control("chkPickupMinors")
    Set chkPickupMajors = g_MainView.Control("chkPickupMajors")
    Set chkUnknownScrolls = g_MainView.Control("chkUnknownScrolls")
    
    Set chkLootFellowshipKills = g_MainView.Control("chkLootFellowshipKills")
    Set lstCorpse = g_MainView.Control("lstCorpse")
    Set chkFilterCorpses = g_MainView.Control("chkFilterCorpses")
    Set chkLootGround = g_MainView.Control("chkLootGround")
    Set chkLootForcePeaceMode = g_MainView.Control("chkLootForcePeaceMode")
    Set chkSendRares = g_MainView.Control("chkSendRares")
    Set chkSendName = g_MainView.Control("chkSendName")
    Set chkBoostLootPriority = g_MainView.Control("chkBoostLootPriority")
    
    Set txtAddCorpse = g_MainView.Control("txtAddCorpse")
    Set btnAddCorpse = g_MainView.Control("btnAddCorpse")
    Set btnCorpseSetSelection = g_MainView.Control("btnCorpseSetSelection")
  
  
    Set chkLootSalvages = g_MainView.Control("chkLootSalvages")
    Set chSalvageMinWork = g_MainView.Control("chSalvageMinWork")
    Set chSalvageSortWork = g_MainView.Control("chSalvageSortWork")
    Set lstLootSalvages = g_MainView.Control("lstLootSalvages")
    Set chSalvage = g_MainView.Control("chSalvage")
    Set cmdAddSalvage = g_MainView.Control("cmdAddSalvage")
    
    Set chkLootArmors = g_MainView.Control("chkLootArmors")
    Set lstLootArmors = g_MainView.Control("lstLootArmors")
    Set chArmorType = g_MainView.Control("chArmorType")
    Set chCoverage = g_MainView.Control("chCoverage")
    Set chMaxArmorWork = g_MainView.Control("chMaxArmorWork")
    Set txtArmorMinAL = g_MainView.Control("txtArmorMinAL")
    Set txtArmorMaxValue = g_MainView.Control("txtArmorMaxValue")
    Set cmdAddArmor = g_MainView.Control("cmdAddArmor")
    
    Set chkLootWeapons = g_MainView.Control("chkLootWeapons")
    Set lstLootWeapons = g_MainView.Control("lstLootWeapons")
    Set chWeaponMaxWork = g_MainView.Control("chWeaponMaxWork")
    Set txtWeaponMinDef = g_MainView.Control("txtWeaponMinDef")
    Set chMaxWieldReq = g_MainView.Control("chMaxWieldReq")
    Set txtWeaponMaxValue = g_MainView.Control("txtWeaponMaxValue")
    Set cmdAddWeapon = g_MainView.Control("cmdAddWeapon")
    
    Set nbWeaponKind = g_MainView.Control("nbWeaponKind")
        Set chMeleeWeapons = g_MainView.Control("chMeleeWeapons")
        Set chWeaponElement = g_MainView.Control("chWeaponElement")
        Set txtWeaponMinLowDmg = g_MainView.Control("txtWeaponMinLowDmg")
        Set txtWeaponMinHighDmg = g_MainView.Control("txtWeaponMinHighDmg")
        Set txtWeaponMinAtt = g_MainView.Control("txtWeaponMinAtt")
    
        Set chMissileWeapons = g_MainView.Control("chMissileWeapons")
        Set chMissileWeaponElement = g_MainView.Control("chMissileWeaponElement")
        Set txtWeaponMinMod = g_MainView.Control("txtWeaponMinMod")
        Set chMinDamageBonus = g_MainView.Control("chMinDamageBonus")
    
    Set chkLootWands = g_MainView.Control("chkLootWands")
    Set lstLootWands = g_MainView.Control("lstLootWands")
    Set chWandMinManaC = g_MainView.Control("chWandMinManaC")
    Set chWandMinMeleeDef = g_MainView.Control("chWandMinMeleeDef")
    Set chWandMinMagicDef = g_MainView.Control("chWandMinMagicDef")
    Set chWandMinPvMBonus = g_MainView.Control("chWandMinPvMBonus")
    Set chWandMaxWieldReq = g_MainView.Control("chWandMaxWieldReq")
    Set chWandMaxWork = g_MainView.Control("chWandMaxWork")
    Set cmdAddWand = g_MainView.Control("cmdAddWand")
    
    Set lstExceptions = g_MainView.Control("lstExceptions")
    Set cmdAddException = g_MainView.Control("cmdAddException")
    
    Set chkSalvageFrequency = g_MainView.Control("chkSalvageFrequency")
    Set txtSalvageFrequency = g_MainView.Control("txtSalvageFrequency")
    Set chkEnableMultipleSalvage = g_MainView.Control("chkEnableMultipleSalvage")
    
    Set chkMinorWard = g_MainView.Control("chkMinorWard")
    Set chkMinorAttribute = g_MainView.Control("chkMinorAttribute")
    Set chkMinorMastery = g_MainView.Control("chkMinorMastery")
    Set chkMajorIgnoreBane = g_MainView.Control("chkMajorIgnoreBane")
    
    'Set chkEnableManaStoneRecharge = g_MainView.Control("chkEnableManaStoneRecharge")
    'Set chStoneType = g_MainView.Control("chStoneType")
    'Set txtMinMana = g_MainView.Control("txtMinMana")
    'Set txtMaxStone = g_MainView.Control("txtMaxStone")
    'Set txtMaxCharge = g_MainView.Control("txtMaxCharge")
    
    Init = True
    MyDebug INTERFACE_NAME & ".Init() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Init = False
    PrintErrorMessage INTERFACE_NAME & ".Init - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function LoadControlsValue() As Boolean
On Error GoTo Error_Handler
    
    MyDebug "[" & INTERFACE_NAME & "] Loading controls settings"
    
    '***************************************************************************
    
    With g_Settings.Profile.Config
        chkEnableLooting.Checked = .GetValue("chkEnableLooting", True)
        chkEnableSalvage.Checked = .GetValue("chkEnableSalvage", True)
        chkSalvageFrequency.Checked = .GetValue("chkSalvageFrequency", True)
        txtSalvageFrequency.Text = .GetValue("txtSalvageFrequency", 10)
        chkEnableMultipleSalvage.Checked = .GetValue("chkEnableMultipleSalvage", True)
    End With
    
    With g_Settings.Profile.LootCfg
        txtPickupRange.Text = .GetValue("txtPickupRange", 4)
        
        chkLootAny.Checked = .GetValue("chkLootAny", False)
        chkPickupValuable.Checked = .GetValue("chkPickupValuable", True)
        txtPickupMinValue.Text = .GetValue("txtPickupMinValue", 10000)
        chkBurdenRatio.Checked = .GetValue("chkBurdenRatio", False)
        txtBurdenRatio.Text = .GetValue("txtBurdenRatio", 10)
        chkPickupMinors.Checked = .GetValue("chkPickupMinors", True)
        chkPickupMajors.Checked = .GetValue("chkPickupMajors", True)
        chkUnknownScrolls.Checked = .GetValue("chkUnknownScrolls", False)
        
        txtHighMana.Text = .GetValue("txtHighMana", 2000)
        
        chkLootFellowshipKills.Checked = .GetValue("chkLootFellowshipKills", False)
        chkFilterCorpses.Checked = .GetValue("chkFilterCorpses", False)
        chkLootGround.Checked = .GetValue("chkLootGround", False)
        chkLootForcePeaceMode.Checked = .GetValue("chkLootForcePeaceMode", True)
        chkSendRares.Checked = .GetValue("chkSendRares", True)
        chkSendName.Checked = .GetValue("chkSendName", False)
        chkBoostLootPriority.Checked = .GetValue("chkBoostLootPriority", False)
        
        'Filters
        chkLootSalvages.Checked = .GetValue("chkLootSalvages", False)
        chSalvageMinWork.Selected = 0
        chSalvageSortWork.Selected = .GetValue("chSalvageSortWork", 5)
        chSalvage.Selected = 0
        
        chkLootArmors.Checked = .GetValue("chkLootArmors", False)
        chArmorType.Selected = 0
        chCoverage.Selected = 0
        chMaxArmorWork.Selected = 0
        txtArmorMinAL.Text = 200
        txtArmorMaxValue.Text = "99999"
       
        chkLootWeapons.Checked = .GetValue("chkLootWeapons", False)
        chWeaponMaxWork.Selected = 0
        chMaxWieldReq.Selected = 0
        txtWeaponMinDef.Text = "0"
        txtWeaponMaxValue.Text = "99999"
        
        chMeleeWeapons.Selected = 0
        chWeaponElement.Selected = 0
        txtWeaponMinLowDmg.Text = "0"
        txtWeaponMinHighDmg.Text = "1"
        txtWeaponMinAtt.Text = "0"
        
        chMissileWeapons.Selected = 0
        chMissileWeaponElement.Selected = 0
        txtWeaponMinMod.Text = "0"
        chMinDamageBonus.Selected = 0
        
        chkLootWands.Checked = .GetValue("chkLootWands", False)
        chWandMinManaC.Selected = 0
        chWandMaxWork.Selected = 0
        chWandMinMeleeDef.Selected = 0
        chWandMinMagicDef.Selected = 0
        chWandMinPvMBonus.Selected = 0
        chWandMaxWieldReq.Selected = 0
        
        chkMinorWard.Checked = .GetValue("chkMinorWard", False)
        chkMinorAttribute.Checked = .GetValue("chkMinorAttribute", False)
        chkMinorMastery.Checked = .GetValue("chkMinorMastery", False)
        chkMajorIgnoreBane.Checked = .GetValue("chkMajorIgnoreBane", False)
        
        'chkEnableManaStoneRecharge.Checked = .GetValue("chkEnableManaStoneRecharge", True)
        'chStoneType.Selected = 0
        'txtMinMana.Text = .GetValue("txtMinMana", 1000)
        'txtMaxStone.Text = .GetValue("txtMaxStone", 3)
        'txtMaxCharge.Text = .GetValue("txtMaxCharge", 3)
    
    End With
    
    Call LoadItemsList(lstPickup, GetPickupListPath)
    'Call LoadItemsList(lstIgnore, GetIgnoreListPath)
    'Call LoadItemsList(lstCorpse, GetCorpseListPath)
    
    'Load Filters Lists
    Call LoadArmorControls
    Call LoadWeaponControls
    Call LoadWandControls
    Call LoadSalvageControls
    
    'Load Exceptions List
    Call UpdateExceptionList
    
    '***************************************************************************

    
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
        .SaveCheckbox chkEnableLooting, "chkEnableLooting"
        .SaveCheckbox chkEnableSalvage, "chkEnableSalvage"
        
        .SaveCheckbox chkEnableMultipleSalvage, "chkEnableMultipleSalvage"
        .SaveCheckbox chkSalvageFrequency, "chkSalvageFrequency"
        .SaveTextbox txtSalvageFrequency, "txtSalvageFrequency"
    End With
    
    With g_Settings.Profile.LootCfg
        .SaveTextbox txtPickupRange, "txtPickupRange"
        .SaveCheckbox chkLootFellowshipKills, "chkLootFellowshipKills"
        .SaveCheckbox chkFilterCorpses, "chkFilterCorpses"
        .SaveCheckbox chkLootGround, "chkLootGround"
        .SaveCheckbox chkLootForcePeaceMode, "chkLootForcePeaceMode"
        .SaveCheckbox chkSendRares, "chkSendRares"
        .SaveCheckbox chkSendName, "chkSendName"
        .SaveCheckbox chkBoostLootPriority, "chkBoostLootPriority"
        
        .SaveCheckbox chkLootAny, "chkLootAny"
        .SaveCheckbox chkPickupValuable, "chkPickupValuable"
        .SaveTextbox txtPickupMinValue, "txtPickupMinValue"
        .SaveCheckbox chkBurdenRatio, "chkBurdenRatio"
        .SaveTextbox txtBurdenRatio, "txtBurdenRatio"
        .SaveCheckbox chkPickupMinors, "chkPickupMinors"
        .SaveCheckbox chkPickupMajors, "chkPickupMajors"
        .SaveCheckbox chkUnknownScrolls, "chkUnknownScrolls"
        
        .SaveTextbox txtHighMana, "txtHighMana"
        
        'Filters
        .SaveCheckbox chkLootSalvages, "chkLootSalvages"
        .SaveChoice chSalvageMinWork, "chSalvageMinWork"
        .SaveChoice chSalvageSortWork, "chSalvageSortWork"
        
        .SaveCheckbox chkLootArmors, "chkLootArmors"
        
        .SaveCheckbox chkLootWeapons, "chkLootWeapons"
                
        .SaveCheckbox chkLootWands, "chkLootWands"
                
        .SaveCheckbox chkMinorWard, "chkMinorWard"
        .SaveCheckbox chkMinorAttribute, "chkMinorAttribute"
        .SaveCheckbox chkMinorMastery, "chkMinorMastery"
        .SaveCheckbox chkMajorIgnoreBane, "chkMajorIgnoreBane"
        
        '.SaveCheckbox chkEnableManaStoneRecharge, "chkEnableManaStoneRecharge"
        '.SaveChoice chStoneType, "chStoneType"
        '.SaveTextbox txtMinMana, "txtMinMana"
        '.SaveTextbox txtMaxStone, "txtMaxStone"
        '.SaveTextbox txtMaxCharge, "txtMaxCharge"
    End With
    
    'Save the items lists
    Call SaveLists
    
    'Save the item exceptions list
    Call g_Data.Exceptions.SaveToFile
    
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

    Set chkEnableLooting = Nothing
    Set chkEnableSalvage = Nothing
    Set txtPickupRange = Nothing

    Set txtAddItemToPickup = Nothing
    Set btnAddItemToPickup = Nothing
    Set txtNumToPickup = Nothing
    Set btnPickupSetSelection = Nothing
    Set lstPickup = Nothing
    
    Set txtAddItemToIgnore = Nothing
    Set btnAddItemToIgnore = Nothing
    Set btnIgnoreSetSelection = Nothing
    Set lstIgnore = Nothing

    Set chkLootAny = Nothing
    Set chkPickupValuable = Nothing
    Set txtPickupMinValue = Nothing
    Set chkBurdenRatio = Nothing
    Set txtBurdenRatio = Nothing
    Set chkPickupMinors = Nothing
    Set chkPickupMajors = Nothing
    Set chkUnknownScrolls = Nothing
    
    Set txtHighMana = Nothing

    Set chkLootFellowshipKills = Nothing
    Set lstCorpse = Nothing
    Set chkFilterCorpses = Nothing
    Set chkLootGround = Nothing
    Set chkLootForcePeaceMode = Nothing
    Set chkSendRares = Nothing
    Set chkSendName = Nothing
    Set chkBoostLootPriority = Nothing
    
    Set txtAddCorpse = Nothing
    Set btnAddCorpse = Nothing
    Set btnCorpseSetSelection = Nothing
    
    Set chkLootSalvages = Nothing
    Set chSalvageMinWork = Nothing
    Set chSalvageSortWork = Nothing
    
    Set lstLootSalvages = Nothing
    Set chSalvage = Nothing
    Set cmdAddSalvage = Nothing
    
    Set chkLootArmors = Nothing
    Set lstLootArmors = Nothing
    Set chArmorType = Nothing
    Set chCoverage = Nothing
    Set chMaxArmorWork = Nothing
    Set txtArmorMinAL = Nothing
    Set txtArmorMaxValue = Nothing
    Set cmdAddArmor = Nothing

       
    Set chkLootWeapons = Nothing
    Set lstLootWeapons = Nothing
    Set chWeaponMaxWork = Nothing
    Set txtWeaponMinAtt = Nothing
    Set txtWeaponMinDef = Nothing
    Set chMaxWieldReq = Nothing
    Set txtWeaponMaxValue = Nothing
    Set cmdAddWeapon = Nothing
    Set nbWeaponKind = Nothing
    Set chMeleeWeapons = Nothing
    Set chWeaponElement = Nothing
    Set txtWeaponMinLowDmg = Nothing
    Set txtWeaponMinHighDmg = Nothing
    Set chMissileWeapons = Nothing
    Set chMissileWeaponElement = Nothing
    Set txtWeaponMinMod = Nothing
    Set chMinDamageBonus = Nothing
    
    Set chkLootWands = Nothing
    Set lstLootWands = Nothing
    Set chWandMinManaC = Nothing
    Set chWandMinMeleeDef = Nothing
    Set chWandMinMagicDef = Nothing
    Set chWandMinPvMBonus = Nothing
    Set chWandMaxWieldReq = Nothing
    Set chWandMaxWork = Nothing
    Set cmdAddWand = Nothing
    
    Set lstExceptions = Nothing
    Set cmdAddException = Nothing
    
    Set chkEnableMultipleSalvage = Nothing
    Set chkSalvageFrequency = Nothing
    Set txtSalvageFrequency = Nothing
    
    Set chkMinorWard = Nothing
    Set chkMinorAttribute = Nothing
    Set chkMinorMastery = Nothing
    Set chkMajorIgnoreBane = Nothing
    
    'Set chkEnableManaStoneRecharge = Nothing
    'Set chStoneType = Nothing
    'Set txtMinMana = Nothing
    'Set txtMaxStone = Nothing
    'Set txtMaxCharge = Nothing
    
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

'#############################################################################################

Public Function GetLootProfilePath() As String
    GetLootProfilePath = g_Settings.GetDataFolder & "\" & PATH_PROFILE_LOOT & "\" & g_Settings.Profile.LootProfileName
End Function

Private Function GetPickupListPath() As String
    GetPickupListPath = GetLootProfilePath & "\" & FILE_ITEMS_TO_PICKUP_LIST
End Function

Private Function GetIgnoreListPath() As String
    GetIgnoreListPath = GetLootProfilePath & "\" & FILE_ITEMS_TO_IGNORE_LIST
End Function

Private Function GetCorpseListPath() As String
    GetCorpseListPath = GetLootProfilePath & "\" & FILE_CORPSES_TO_LOOT_LIST
End Function

'returns -1 if couldn't find item
Public Function FindItemInList(ByRef lstControl As DecalControls.list, FullItemName As String) As Long
On Error GoTo Error_Handler

    Dim i As Integer
    Dim curItemName As String
    Dim bExactMatch As Boolean

    'default return value
    FindItemInList = -1
    
    For i = 0 To lstControl.Count - 1
        curItemName = lstControl.Data(COL_ITEM_NAME, i)
        bExactMatch = lstControl.Data(COL_EXACT_MATCH, i)
        
        'try to find the full item name in the list before a partial name entry
        If (bExactMatch And SameText(LCase(curItemName), LCase(FullItemName))) Then
            FindItemInList = i
            Exit Function
        ElseIf ((Not bExactMatch) And InStr(LCase(FullItemName), LCase(curItemName))) Then
            FindItemInList = i
        End If
        
    Next i
    
Fin:
    Exit Function
Error_Handler:
    FindItemInList = -1
    PrintErrorMessage INTERFACE_NAME & ".ItemInList : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Function

Public Function ItemInPickupList(FullItemName As String, Optional bMustBeChecked As Boolean = False) As Boolean
    If bMustBeChecked Then
        ItemInPickupList = ItemCheckedInList(lstPickup, FullItemName)
    Else
        ItemInPickupList = (FindItemInList(lstPickup, FullItemName) >= 0)
    End If
End Function

Public Function ItemInIgnoreList(FullItemName As String, Optional bMustBeChecked As Boolean = False) As Boolean
    If bMustBeChecked Then
        ItemInIgnoreList = ItemCheckedInList(lstIgnore, FullItemName)
    Else
        ItemInIgnoreList = (FindItemInList(lstIgnore, FullItemName) >= 0)
    End If
End Function

Public Function ItemInCorpseList(CorpseName As String, Optional bMustBeChecked As Boolean = False) As Boolean
    If bMustBeChecked Then
        ItemInCorpseList = ItemCheckedInList(lstCorpse, CorpseName)
    Else
        ItemInCorpseList = (FindItemInList(lstCorpse, CorpseName) >= 0)
    End If
End Function

'Returns true if a match was found in the list
'Stops at the first match
Public Function ItemCheckedInList(ByRef lstControl As DecalControls.list, FullItemName As String) As Boolean
On Error GoTo Error_Handler

    Dim i As Integer
    
    i = FindItemInList(lstControl, FullItemName)
    If i >= 0 Then
        ItemCheckedInList = lstControl.Data(COL_ENABLED, i)
    Else
        ItemCheckedInList = False
    End If
    
Fin:
    Exit Function
Error_Handler:
    ItemCheckedInList = False
    PrintErrorMessage INTERFACE_NAME & ".ItemCheckedInList : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Function

'Finds the number of items in our inventory
Public Function ItemCheckCount(ByRef lstControl As DecalControls.list, FullItemName As String) As Boolean
On Error GoTo Error_Handler

    Dim i As Integer
    Dim iNumCount As Integer
    
    i = FindItemInList(lstControl, FullItemName)
    If i >= 0 Then
        iNumCount = CInt(lstControl.Data(COL_NUM_PICKUP, i))
    Else
        iNumCount = 0
    End If
    
    If iNumCount > 0 Then
        'Compare to # we currently have in our Inventory
        If iNumCount > g_Objects.Items.InvCntByName(FullItemName) Then
            ItemCheckCount = True
        Else
            ItemCheckCount = False
        End If
    Else
        ItemCheckCount = True
    End If
    
Fin:
    Exit Function
Error_Handler:
    ItemCheckCount = True
    PrintErrorMessage INTERFACE_NAME & ".ItemCheckCount : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Function

Public Function ItemIsPickable(FullItemName As String, Optional bExactMatch As Boolean = True) As Boolean
On Error GoTo Error_Handler

    Dim i As Integer
    
    'First make sure this item is not on our ignore list
    If ItemCheckedInList(lstIgnore, FullItemName) Then
        'MyDebug "ItemIsPickable(" & FullItemName & ") - Item on ignore list. Ignoring."
        ItemIsPickable = False
        Exit Function
    End If
    
    'Then check if it's on our pickup list
    If ItemCheckedInList(lstPickup, FullItemName) Then
        If ItemCheckCount(lstPickup, FullItemName) Then
            ItemIsPickable = True
        Else
            ItemIsPickable = False
        End If
    Else
        ItemIsPickable = False
    End If

Fin:
    Exit Function
Error_Handler:
    ItemIsPickable = False
    PrintErrorMessage INTERFACE_NAME & ".ItemIsPickable : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Function

Public Function AddItemToList(ByRef lstControl As DecalControls.list, ByVal sItemName As String, Optional ByVal bEnabled As Boolean = True, Optional ByVal bExactMatch As Boolean = True, Optional ByVal iNumToPickup As Integer = 0) As Boolean
On Error GoTo Error_Handler
    
    Dim i As Integer, j As Integer
    
    Dim curItemName As String
    Dim curExactMatch As Boolean
    
    'MyDebug INTERFACE_NAME & ".AddItemToList: " & sItemName
    
    'First check if it's not already in the list
    For i = 0 To lstControl.Count - 1
        curItemName = lstControl.Data(COL_ITEM_NAME, i)
        curExactMatch = lstControl.Data(COL_EXACT_MATCH, i)
        
        If SameText(curItemName, sItemName) Then
            PrintMessage sItemName & " already in the list."
            AddItemToList = False
            Exit Function
        
        'ex: "Steel" is in the list (curExactMatch = false)
        'and we want to add "Steel Platemail" (bExactMatch = false)
        '-> error message
        ElseIf (Not curExactMatch) _
        And (InStr(1, sItemName, curItemName) > 0) Then
            PrintMessage "NOTICE : You've already added " & curItemName & " items type to the list - " & sItemName & " belongs to this category."
        End If
    Next i
    
    'find the good insertion position in the list of item names
    i = -1
    If lstControl.Count > 0 Then
        For j = 0 To lstControl.Count - 1
            curItemName = lstControl.Data(COL_ITEM_NAME, j)
            If StrComp(curItemName, sItemName, vbTextCompare) > 0 Then
                i = j
                Exit For
            End If
        Next j
    End If
    
    If i = -1 Then
        i = lstControl.AddRow
    Else
        Call lstControl.InsertRow(i)
    End If

    With lstControl
         curItemName = Trim(sItemName)
        .Data(COL_ENABLED, i) = bEnabled
        .Data(COL_ITEM_NAME, i) = curItemName
        .Data(COL_EXACT_MATCH, i) = bExactMatch
        If iNumToPickup > 0 Then
            .Data(COL_NUM_PICKUP, i) = iNumToPickup
        Else
            'Do Nothing!
            '.Data(COL_NUM_PICKUP, i) = ""
        End If
    End With
    
    AddItemToList = True
    
Fin:
    Exit Function
Error_Handler:
    AddItemToList = False
    PrintErrorMessage INTERFACE_NAME & ".AddItemToList : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
    
End Function

'Save the items list
Public Function SaveItemsList(ByRef lstControl As DecalControls.list, ByVal sFullPath As String) As Boolean
On Error GoTo Error_Handler
    
    Dim bRet As Boolean
    Dim db As New DataFile
    Dim dat As clsDataEntry
    
    Dim i As Integer
    
    MyDebug INTERFACE_NAME & ".SaveItemsList - Saving list to " & sFullPath
    MyDebug INTERFACE_NAME & ".SaveItemsList - CountCols: " & lstControl.CountCols
    
    For i = 0 To lstControl.Count - 1
        Set dat = New clsDataEntry
        dat.AddParam TAG_ENABLED, BoolToInteger(lstControl.Data(COL_ENABLED, i))
        dat.AddParam TAG_ITEM_NAME, lstControl.Data(COL_ITEM_NAME, i)
        dat.AddParam TAG_EXACT_MATCH, lstControl.Data(COL_EXACT_MATCH, i)
        If lstControl.CountCols > 3 Then
            If CInt(lstControl.Data(COL_NUM_PICKUP, i)) > 0 Then
                dat.AddParam TAG_NUM_PICKUP, lstControl.Data(COL_NUM_PICKUP, i)
            Else
                dat.AddParam TAG_NUM_PICKUP, 0
            End If
        End If
        
        If Not db.AddData(dat) Then
            PrintErrorMessage INTERFACE_NAME & ".SaveItemsList - db.AddData(dat) failed"
        End If
    Next i
    
    bRet = db.save(sFullPath)
    
    If Not bRet Then
        PrintErrorMessage INTERFACE_NAME & ".SaveItemsList - Failed to save list to " & sFullPath
    End If
    
Fin:
    Set db = Nothing
    Set dat = Nothing
    SaveItemsList = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage INTERFACE_NAME & ".SaveItemsList(" & sFullPath & ") - " & Err.Description
    Resume Fin
End Function

Public Function LoadItemsList(ByRef lstControl As DecalControls.list, ByVal sFullPath As String) As Boolean
On Error GoTo Error_Handler
    
    Dim bRet As Boolean
    Dim db As New DataFile
    Dim dat As clsDataEntry
    
    Dim i As Integer
    
    MyDebug INTERFACE_NAME & ".LoadItemsList - Loading list from " & sFullPath
    
    'first reset the list
    Call lstControl.Clear
    
    'Load the data
    If db.Load(sFullPath) Then
    
        Dim bEnable As Boolean
        Dim bExactMatch As Boolean
        Dim sItemName As String
        Dim iNumToPickup As Integer
        
        For Each dat In db
            If dat.ParamExist(TAG_ITEM_NAME) Then
                sItemName = dat.Param(TAG_ITEM_NAME)
                
                If dat.ParamExist(TAG_EXACT_MATCH) Then
                    bExactMatch = dat.Param(TAG_EXACT_MATCH)
                Else
                    bExactMatch = True
                End If
                
                If dat.ParamExist(TAG_ENABLED) Then
                    bEnable = dat.Param(TAG_ENABLED)
                Else
                    bEnable = False
                End If
                
                iNumToPickup = 0
                If dat.ParamExist(TAG_NUM_PICKUP) Then
                    If dat.Param(TAG_NUM_PICKUP) <> "" Then
                        iNumToPickup = CInt(dat.Param(TAG_NUM_PICKUP))
                    End If
                End If
                
                Call AddItemToList(lstControl, Trim(sItemName), bEnable, bExactMatch, iNumToPickup)
            End If
        Next dat
    Else
        PrintErrorMessage INTERFACE_NAME & ".LoadItemsList - Failed to load " & sFullPath
    End If
    
Fin:
    Set db = Nothing
    Set dat = Nothing
    LoadItemsList = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage INTERFACE_NAME & ".LoadItemsList(" & sFullPath & ") - " & Err.Description
    Resume Fin
End Function

Public Sub SaveLists()
    Call SaveItemsList(lstPickup, GetPickupListPath)
    Call SaveItemsList(lstIgnore, GetIgnoreListPath)
    Call SaveItemsList(lstCorpse, GetCorpseListPath)
End Sub

'#############################################################################################

Private Sub btnAddCorpse_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    If AddItemToList(lstCorpse, txtAddCorpse.Text, True, True) Then
        Call SaveItemsList(lstCorpse, GetCorpseListPath)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage INTERFACE_NAME & ".btnAddCorpse_Accepted : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Sub

Private Sub btnAddItemToIgnore_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    If AddItemToList(lstIgnore, txtAddItemToIgnore.Text, True, True) Then
        Call SaveItemsList(lstIgnore, GetIgnoreListPath)
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage INTERFACE_NAME & ".bntAddItemToIgnore_Accepted : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Sub


Private Sub btnAddItemToPickup_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    Dim ItemName As String
    ItemName = txtAddItemToPickup.Text
    
    If ItemCheckedInList(lstIgnore, ItemName) Then
        PrintMessage "Impossible to add " & ItemName & " to the Pickup list : " & ItemName & " is on the Ignored items list. Remove it from there first."
    Else
        If AddItemToList(lstPickup, txtAddItemToPickup.Text, True, True, g_Data.NumToPickup) Then
            Call SaveItemsList(lstPickup, GetPickupListPath)
        End If
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage INTERFACE_NAME & ".bntAddItemToPickup_Accepted : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
End Sub

Private Sub btnCorpseSetSelection_Accepted(ByVal nID As Long)
    Call GetCurrentSelection(txtAddCorpse)
End Sub

Private Sub btnIgnoreSetSelection_Accepted(ByVal nID As Long)
    Call GetCurrentSelection(txtAddItemToIgnore)
End Sub

Private Sub btnPickupSetSelection_Accepted(ByVal nID As Long)
    Call GetCurrentSelection(txtAddItemToPickup)
End Sub

Private Sub lstCorpse_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Call CommonListChange(lstCorpse, nX, nY)
End Sub

Private Sub lstIgnore_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Call CommonListChange(lstIgnore, nX, nY)
End Sub

Private Sub lstLootArmors_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
On Error GoTo ErrorHandler

    Select Case nX
        Case COL_ARMOR_TYPE
            Call RemoveRowFromFilterList(nY, lstLootArmors, FILTER_ARMOR)
        Case COL_ARMOR_ENABLE
            Call EnableRowFromFilterList(nY, lstLootArmors, FILTER_ARMOR)
        
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".lstLootArmors_Change - " & Err.Description
    Resume Fin
End Sub


Private Sub lstLootSalvages_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
On Error GoTo ErrorHandler

    Select Case nX
        Case COL_SALVAGE_NAME
            Call RemoveRowFromFilterList(nY, lstLootSalvages, FILTER_SALVAGE)
        Case COL_SALVAGE_ENABLE
            Call EnableRowFromFilterList(nY, lstLootSalvages, FILTER_SALVAGE)
        Case COL_SALVAGE_COMBINE
            Call EnableSalvageCombine(nY)
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".lstLootSalvages_Change - " & Err.Description
    Resume Fin
End Sub


Private Sub lstLootWeapons_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
On Error GoTo ErrorHandler

    Select Case nX
        Case COL_WEAPON_TYPE
            Call RemoveRowFromFilterList(nY, lstLootWeapons, FILTER_WEAPON)
        Case COL_WEAPON_ENABLE
            Call EnableRowFromFilterList(nY, lstLootWeapons, FILTER_WEAPON)
            
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".lstLootWeapons_Change - " & Err.Description
    Resume Fin
End Sub

Private Sub lstLootWands_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
On Error GoTo ErrorHandler

    Select Case nX
        Case COL_WAND_TYPE
            Call RemoveRowFromFilterList(nY, lstLootWands, FILTER_WAND)
        Case COL_WAND_ENABLE
            Call EnableRowFromFilterList(nY, lstLootWands, FILTER_WAND)
            
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".lstLootWands_Change - " & Err.Description
    Resume Fin
End Sub

Private Sub lstPickup_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Call CommonListChange(lstPickup, nX, nY)
End Sub

'###################################################################################################

Private Sub EnableSalvageCombine(ByVal lRowIndex As Long)
On Error GoTo ErrorHandler

    If (lRowIndex >= 0) And (lRowIndex < lstLootSalvages.Count) Then
        Dim lFilterId As Long
        Dim colFilter As Collection
        
        Set colFilter = g_Data.LootFilters.SalvageFilters
        
        If m_dicRowToFilterId(FILTER_SALVAGE).Exists(lRowIndex) Then
            
            lFilterId = m_dicRowToFilterId(FILTER_SALVAGE).Item(lRowIndex)
            
            If Valid(colFilter) Then
                If g_Data.LootFilters.SalvageCombineFilter(colFilter, lFilterId) Then
                    MyDebug "EnableSalvageCombine - Filter #" & lFilterId & " salvage combin toggled"
                Else
                    MyDebug "EnableSalvageCombine - FAILED to toggle Filter #" & lFilterId & " from Filter collection [" & FILTER_SALVAGE & "]"
                    MyDebug "Available Filters in this collection : "
                    Dim oFilter As clsLootFilter
                    For Each oFilter In colFilter
                        MyDebug "... FilterId #" & oFilter.FilterId
                    Next oFilter
                End If
            End If
            
        Else
            MyDebug "EnableSalvageCombine - RowIndex " & lRowIndex & " doesnt exist in dicRowToFilter(" & FILTER_SALVAGE & ")"
        End If
                
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".EnableSalvageCombine - " & Err.Description
    Resume Fin
End Sub


Private Sub EnableRowFromFilterList(ByVal lRowIndex As Long, lstFilter As DecalControls.list, ByVal lFilterKind As Long)
On Error GoTo ErrorHandler

    If (lRowIndex >= 0) And (lRowIndex < lstFilter.Count) Then
        Dim lFilterId As Long
        Dim colFilter As Collection
        
        If m_dicRowToFilterId(lFilterKind).Exists(lRowIndex) Then
            lFilterId = m_dicRowToFilterId(lFilterKind).Item(lRowIndex)
            
            Select Case lFilterKind
                Case FILTER_ARMOR
                    'MyDebug "Armor Filter"
                    Set colFilter = g_Data.LootFilters.ArmorFilters
                    
                Case FILTER_WEAPON
                    'MyDebug "Weapon Filter"
                    Set colFilter = g_Data.LootFilters.WeaponFilters
                
                Case FILTER_WAND
                    'MyDebug "Wand Filter"
                    Set colFilter = g_Data.LootFilters.WandFilters
                    
                Case FILTER_SALVAGE
                    'MyDebug "Salvage Fitler"
                    Set colFilter = g_Data.LootFilters.SalvageFilters
                    
                Case Else
                    Set colFilter = Nothing
            End Select
            
            If Valid(colFilter) Then
                If g_Data.LootFilters.EnableFilter(colFilter, lFilterId) Then
                    MyDebug "EnableRowFromFilterList - Filter #" & lFilterId & " toggled"
                Else
                    MyDebug "EnableRowFromFilterList - FAILED to toggle Filter #" & lFilterId & " from Filter collection [" & lFilterKind & "]"
                    MyDebug "Available Filters in this collection : "
                    Dim oFilter As clsLootFilter
                    For Each oFilter In colFilter
                        MyDebug "... FilterId #" & oFilter.FilterId
                    Next oFilter
                End If
            End If
            
            'Update Row<>FilterId dictionary
            'Call m_dicRowToFilterId(lFilterKind).Remove(lRowIndex)
            
        Else
            MyDebug "EnableRowFromFilterList - RowIndex " & lRowIndex & " doesnt exist in dicRowToFilter(" & lFilterKind & ")"
        End If
                
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".EnableRowFromFilterList - " & Err.Description
    Resume Fin
End Sub


Private Sub RemoveRowFromFilterList(ByVal lRowIndex As Long, lstFilter As DecalControls.list, ByVal lFilterKind As Long)
On Error GoTo ErrorHandler

    If (lRowIndex >= 0) And (lRowIndex < lstFilter.Count) Then
        Dim lFilterId As Long
        Dim colFilter As Collection
        
        If m_dicRowToFilterId(lFilterKind).Exists(lRowIndex) Then
            lFilterId = m_dicRowToFilterId(lFilterKind).Item(lRowIndex)
            
            Select Case lFilterKind
                Case FILTER_ARMOR
                    'MyDebug "Armor Filter"
                    Set colFilter = g_Data.LootFilters.ArmorFilters
                    
                Case FILTER_WEAPON
                    'MyDebug "Weapon Filter"
                    Set colFilter = g_Data.LootFilters.WeaponFilters
                    
                Case FILTER_WAND
                    'MyDebug "Wand Filter"
                    Set colFilter = g_Data.LootFilters.WandFilters
                    
                Case FILTER_SALVAGE
                    'MyDebug "Salvage Fitler"
                    Set colFilter = g_Data.LootFilters.SalvageFilters
                    
                Case Else
                    Set colFilter = Nothing
            End Select
            
            If Valid(colFilter) Then
                If g_Data.LootFilters.RemoveFilter(colFilter, lFilterId) Then
                    MyDebug "RemoveRowFromFilterList - Filter #" & lFilterId & " removed from Filter collection successfully"
                Else
                    MyDebug "RemoveRowFromFilterList - FAILED to remove Filter #" & lFilterId & " from Filter collection [" & lFilterKind & "]"
                    MyDebug "Available Filters in this collection : "
                    Dim oFilter As clsLootFilter
                    For Each oFilter In colFilter
                        MyDebug "... FilterId #" & oFilter.FilterId
                    Next oFilter
                End If
            End If
            
            'Update Row<>FilterId dictionary
            'Call m_dicRowToFilterId(lFilterKind).Remove(lRowIndex)
            
        Else
            MyDebug "RemoveRowFromFilterList - RowIndex " & lRowIndex & " doesnt exist in dicRowToFilter(" & lFilterKind & ")"
        End If
        
        'Reload List
        Select Case lFilterKind
            Case FILTER_ARMOR
                Call LoadArmorList
                
            Case FILTER_WEAPON
                Call LoadWeaponsList
            
            Case FILTER_WAND
                Call LoadWandList
                
            Case FILTER_SALVAGE
                Call LoadSalvageList
                
            Case Else
                Call lstFilter.DeleteRow(lRowIndex)
        End Select
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".RemoveRowFromFilterList - " & Err.Description
    Resume Fin
End Sub

Private Sub LinkRowToFilter(ByVal lRowIndex As Long, ByVal oFilter As clsLootFilter, ByVal lFilterKind As Long)
On Error GoTo ErrorHandler
    
    If Valid(oFilter) Then
        'overwrite if already exist
        Dim dicRow As Dictionary
        Set dicRow = m_dicRowToFilterId(lFilterKind)
        If dicRow.Exists(lRowIndex) Then
            MyDebug "LinkRowToFilter[" & lFilterKind & "] - line #" & lRowIndex & " is already linked to filter #" & dicRow(lRowIndex) & ". Over-writting it by filter #" & oFilter.FilterId
            m_dicRowToFilterId(lFilterKind).Item(lRowIndex) = oFilter.FilterId
        Else
            'MyDebug "LinkRowToFilter[" & lFilterKind & "] Linking RowIndex #" & lRowIndex & " to FilterId #" & oFilter.FilterId
            Call m_dicRowToFilterId(lFilterKind).Add(lRowIndex, oFilter.FilterId)
        End If
        Set dicRow = Nothing
    Else
        PrintErrorMessage INTERFACE_NAME & "LinkRowToFilter[" & lFilterKind & "] (raw:" & lRowIndex & ") - Invalid oFilter"
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LinkRowToFilter(" & lRowIndex & ", " & lFilterKind & ") - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

'--------------- Salvages --------------------
Private Sub LoadSalvageList()
On Error GoTo ErrorHandler
    Dim oFilter As clsLootFilter
    
    Call lstLootSalvages.Clear
    Call m_dicRowToFilterId(FILTER_SALVAGE).RemoveAll
    
    For Each oFilter In g_Data.LootFilters.SalvageFilters
        Call AddSalvageFilterToList(oFilter)
    Next oFilter
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadSalvageList - " & Err.Description
    Resume Fin
End Sub
Private Sub LoadSalvageControls()
On Error GoTo ErrorHandler

    Dim oData As DarksideFilter.DataItem
    
    'Fill the salvage types dropdown
    Call chSalvage.Clear
    For Each oData In g_ACConst.Materials
        Call chSalvage.AddChoice(oData.Val)
    Next oData
    If chSalvage.ChoiceCount > 0 Then
        chSalvage.Selected = 0
    End If
    
    Call LoadSalvageList
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadSalvageControls - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdAddSalvage_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    Dim oFilter As New clsLootFilter
    Dim oCurFilter As clsLootFilter
    
    With oFilter
        .FilterType = FILTER_SALVAGE
        .SalvageMaterial = g_ACConst.Materials.GetKey(chSalvage.Text(chSalvage.Selected))
        .MinWork = Val(chSalvageMinWork.Text(chSalvageMinWork.Selected))
        .FilterEnabled = True
        .SalvageCombine = True
    End With

    For Each oCurFilter In g_Data.LootFilters.SalvageFilters
        If oCurFilter.SalvageMaterial = oFilter.SalvageMaterial Then
            'just update the already existing filter
            Dim i As Long
            
            'Find the list row index to update
            For i = 0 To lstLootSalvages.Count - 1
                If m_dicRowToFilterId(FILTER_SALVAGE).Exists(i) Then
                    If m_dicRowToFilterId(FILTER_SALVAGE).Item(i) = oCurFilter.FilterId Then
                        'update filter's min workmanship
                        oCurFilter.FilterEnabled = oFilter.FilterEnabled
                        lstLootSalvages.Data(COL_SALVAGE_ENABLE, i) = 1
                        oCurFilter.MinWork = oFilter.MinWork
                        lstLootSalvages.Data(COL_SALVAGE_MIN_WORK, i) = oCurFilter.MinWork
                        oCurFilter.SalvageCombine = oFilter.SalvageCombine
                        lstLootSalvages.Data(COL_SALVAGE_COMBINE, i) = 1
                        lstLootSalvages.ScrollPosition = i
                        PrintMessage g_ACConst.Materials(oCurFilter.SalvageMaterial) & " salvage Minimum Workmanship changed to " & oCurFilter.MinWork
                        GoTo Fin
                    End If
                End If
            Next i
        End If
    Next oCurFilter
    
    'If we reach this line, it means we're adding a new filter
    If g_Data.LootFilters.AddSalvageFilter(oFilter) Then
        'filterId updated
        Call AddSalvageFilterToList(oFilter)
    End If
    
Fin:
    Set oFilter = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".cmdAddSalvage_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub AddSalvageFilterToList(oFilter As clsLootFilter)
On Error GoTo ErrorHandler
    
    If Valid(oFilter) Then
        Dim i As Long
        i = lstLootSalvages.AddRow
        Call LinkRowToFilter(i, oFilter, FILTER_SALVAGE)
        With lstLootSalvages
            .Data(COL_SALVAGE_NAME, i) = g_ACConst.Materials(oFilter.SalvageMaterial)
            .Color(COL_SALVAGE_NAME, i) = m_COLOR_DEL_COL
            .Data(COL_SALVAGE_MIN_WORK, i) = oFilter.MinWork
            .ScrollPosition = i
        End With
        If oFilter.FilterEnabled Then
            lstLootSalvages.Data(COL_SALVAGE_ENABLE, i) = 1
        Else
            lstLootSalvages.Data(COL_SALVAGE_ENABLE, i) = 0
        End If
        If oFilter.SalvageCombine Then
            lstLootSalvages.Data(COL_SALVAGE_COMBINE, i) = 1
        Else
            lstLootSalvages.Data(COL_SALVAGE_COMBINE, i) = 0
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".AddSalvageFilterToList - " & Err.Description
    Resume Fin
End Sub

'--------------- Armors --------------------
Private Sub LoadArmorList()
On Error GoTo ErrorHandler
    Dim oFilter As clsLootFilter
    
    Call lstLootArmors.Clear
    Call m_dicRowToFilterId(FILTER_ARMOR).RemoveAll
    
    For Each oFilter In g_Data.LootFilters.ArmorFilters
        Call AddArmorFilterToList(oFilter)
    Next oFilter
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadArmorList - " & Err.Description
    Resume Fin
End Sub

Private Sub LoadArmorControls()
On Error GoTo ErrorHandler

    Dim oData As DarksideFilter.DataItem
    
    'Fill the armor types dropdown
    Call chArmorType.Clear
    For Each oData In g_ACConst.Armors
        Call chArmorType.AddChoice(oData.Val)
    Next oData
    If chArmorType.ChoiceCount > 0 Then
        chArmorType.Selected = 0
    End If
    
    'Fill the armor coverage dropdown
    Call chCoverage.Clear
    For Each oData In g_ACConst.ArmorCoverage
        Call chCoverage.AddChoice(oData.Val)
    Next oData
    If chCoverage.ChoiceCount > 0 Then
        chCoverage.Selected = 0
    End If
    
    Call LoadArmorList
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadArmorControls - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdAddArmor_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    Dim oFilter As New clsLootFilter
    
    With oFilter
        .FilterType = FILTER_ARMOR
        .FilterEnabled = True
        .ArmorType = g_ACConst.Armors.GetKey(chArmorType.Text(chArmorType.Selected))
        .ArmorCoverage = g_ACConst.ArmorCoverage.GetKey(chCoverage.Text(chCoverage.Selected))
        .ArmorMinAL = Val(txtArmorMinAL.Text)
        .MaxWork = Val(chMaxArmorWork.Text(chMaxArmorWork.Selected))
        .MaxValue = Val(txtArmorMaxValue.Text)
    End With
    
    If g_Data.LootFilters.AddArmorFilter(oFilter) Then
        'filterId updated
        Call AddArmorFilterToList(oFilter)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".cmdAddArmor_Accepted - " & Err.Description
    Resume Fin
End Sub



Private Sub AddArmorFilterToList(oFilter As clsLootFilter)
On Error GoTo ErrorHandler
    
    If Valid(oFilter) Then
        Dim i As Long
        i = lstLootArmors.AddRow
        Call LinkRowToFilter(i, oFilter, FILTER_ARMOR)
        With lstLootArmors
            .Data(COL_ARMOR_TYPE, i) = g_ACConst.Armors(oFilter.ArmorType)
            .Color(COL_ARMOR_TYPE, i) = m_COLOR_DEL_COL
            .Data(COL_ARMOR_COVERAGE, i) = g_ACConst.ArmorCoverage(oFilter.ArmorCoverage)
            .Data(COL_ARMOR_MIN_AL, i) = oFilter.ArmorMinAL
            .Data(COL_ARMOR_MAX_WORK, i) = oFilter.MaxWork
            .Data(COL_ARMOR_MAX_VALUE, i) = oFilter.MaxValue
            .ScrollPosition = i
        End With
        If (oFilter.FilterEnabled) Then
            lstLootArmors.Data(COL_ARMOR_ENABLE, i) = 1
        Else
            lstLootArmors.Data(COL_ARMOR_ENABLE, i) = 0
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".AddArmorFilterToList - " & Err.Description
    Resume Fin
End Sub


'--------------- Weapons --------------------
Private Sub LoadWeaponsList()
On Error GoTo ErrorHandler
    Dim oFilter As clsLootFilter
    
    Call lstLootWeapons.Clear
    Call m_dicRowToFilterId(FILTER_WEAPON).RemoveAll
    
    For Each oFilter In g_Data.LootFilters.WeaponFilters
        Call AddWeaponFilterToList(oFilter)
    Next oFilter
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadWeaponsList - " & Err.Description
    Resume Fin
End Sub

Private Sub LoadWeaponControls()
On Error GoTo ErrorHandler
    Dim oData As DarksideFilter.DataItem
    
    'Fill the weapons type dropdown
    Call chMeleeWeapons.Clear
    Call chMissileWeapons.Clear
    
    For Each oData In g_ACConst.WeaponSkills
        If oData.Flag = ITEM_MELEE_WEAPON Then
            Call chMeleeWeapons.AddChoice(oData.Val)
        Else
            Call chMissileWeapons.AddChoice(oData.Val)
        End If
    Next oData
    
    If chMeleeWeapons.ChoiceCount > 0 Then chMeleeWeapons.Selected = 0
    If chMissileWeapons.ChoiceCount > 0 Then chMissileWeapons.Selected = 0
    
    'Fill weapon element dropdown
    Call chWeaponElement.Clear
    For Each oData In g_ACConst.WeaponDamages
        Call chWeaponElement.AddChoice(oData.Val)
    Next oData
    If chWeaponElement.ChoiceCount > 0 Then chWeaponElement.Selected = 0
    
    Call chMissileWeaponElement.Clear
    For Each oData In g_ACConst.WeaponDamages
        Call chMissileWeaponElement.AddChoice(oData.Val)
    Next oData
    If chMissileWeaponElement.ChoiceCount > 0 Then chMissileWeaponElement.Selected = 0
    
    'Fill weapons filters list
    Call LoadWeaponsList
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadWeaponControls - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdAddWeapon_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    Dim oFilter As New clsLootFilter
    
    With oFilter
        .FilterType = FILTER_WEAPON
        .FilterEnabled = True
        
        If nbWeaponKind.ActiveTab = TAB_MELEE_WEAPONS Then
            .WeaponType = g_ACConst.WeaponSkills.GetKey(chMeleeWeapons.Text(chMeleeWeapons.Selected))
            .WeaponMinHighDmg = Val(txtWeaponMinHighDmg.Text)
            .WeaponMinLowDmg = Val(txtWeaponMinLowDmg.Text)
            .WeaponMinAttackBonus = Val(txtWeaponMinAtt.Text)
            .WeaponElement = g_ACConst.WeaponDamages.GetKey(chWeaponElement.Text(chWeaponElement.Selected))
        Else
            .WeaponType = g_ACConst.WeaponSkills.GetKey(chMissileWeapons.Text(chMissileWeapons.Selected))
            .WeaponMinMod = Val(txtWeaponMinMod.Text)
            .WeaponMinBonusDmg = Val(chMinDamageBonus.Text(chMinDamageBonus.Selected))
            .WeaponElement = g_ACConst.WeaponDamages.GetKey(chMissileWeaponElement.Text(chMissileWeaponElement.Selected))
        End If
        
        .WeaponMinDefenseBonus = Val(txtWeaponMinDef.Text)
        .MaxWork = Val(chWeaponMaxWork.Text(chWeaponMaxWork.Selected))
        .WeaponMaxWieldReq = Val(chMaxWieldReq.Text(chMaxWieldReq.Selected))
        .MaxValue = Val(txtWeaponMaxValue.Text)
    End With
    
    If g_Data.LootFilters.AddWeaponFilter(oFilter) Then
        'filterId updated
        Call AddWeaponFilterToList(oFilter)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".cmdAddWeapon_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub AddWeaponFilterToList(oFilter As clsLootFilter)
On Error GoTo ErrorHandler
    
    If Valid(oFilter) Then
        Dim i As Long
        i = lstLootWeapons.AddRow
        Call LinkRowToFilter(i, oFilter, FILTER_WEAPON)
        With lstLootWeapons
            .Data(COL_WEAPON_TYPE, i) = g_ACConst.GetWeaponSkillName(oFilter.WeaponType)    'fixme : convert to string
            .Color(COL_WEAPON_TYPE, i) = m_COLOR_DEL_COL
            
            If g_ACConst.WeaponSkills.Item(oFilter.WeaponType).Flag = ITEM_MISSILE_WEAPON Then
                .Data(COL_WEAPON_DMG, i) = "+" & oFilter.WeaponMinMod & "%"
                .Data(COL_WEAPON_ATT, i) = "+" & oFilter.WeaponMinBonusDmg
            Else    'melee weapons
                .Data(COL_WEAPON_DMG, i) = oFilter.WeaponMinLowDmg & "-" & oFilter.WeaponMinHighDmg
                .Data(COL_WEAPON_ATT, i) = "+" & oFilter.WeaponMinAttackBonus
            End If
            
            .Data(COL_WEAPON_ELEMENT, i) = g_ACConst.WeaponDamages(oFilter.WeaponElement)
            .Data(COL_WEAPON_DEF, i) = "+" & oFilter.WeaponMinDefenseBonus
            .Data(COL_WEAPON_MAX_WORK, i) = oFilter.MaxWork
            .Data(COL_WEAPON_WIELD_REQ, i) = oFilter.WeaponMaxWieldReq
            .Data(COL_WEAPON_MAX_VALUE, i) = oFilter.MaxValue
            .ScrollPosition = i
        End With
        If oFilter.FilterEnabled Then
            lstLootWeapons.Data(COL_WEAPON_ENABLE, i) = 1
        Else
            lstLootWeapons.Data(COL_WEAPON_ENABLE, i) = 0
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".AddWeaponFilterToList - " & Err.Description
    Resume Fin
End Sub

' ----------------  Wands  ----------------------

Private Sub LoadWandList()
On Error GoTo ErrorHandler
    Dim oFilter As clsLootFilter
    
    Call lstLootWands.Clear
    Call m_dicRowToFilterId(FILTER_WAND).RemoveAll
    
    For Each oFilter In g_Data.LootFilters.WandFilters
        Call AddWandFilterToList(oFilter)
    Next oFilter
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadWandList - " & Err.Description
    Resume Fin
End Sub

Private Sub LoadWandControls()
On Error GoTo ErrorHandler

    Dim oData As DarksideFilter.DataItem
    
    Call LoadWandList
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".LoadWandControls - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdAddWand_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    Dim oFilter As New clsLootFilter
    
    With oFilter
        .FilterType = FILTER_WAND
        .FilterEnabled = True
        
        .WandManaC = Val(chWandMinManaC.Text(chWandMinManaC.Selected))
        .WandMeleeD = Val(chWandMinMeleeDef.Text(chWandMinMeleeDef.Selected))
        .WandMagicD = Val(chWandMinMagicDef.Text(chWandMinMagicDef.Selected))
        .WandPVM = Val(chWandMinPvMBonus.Text(chWandMinPvMBonus.Selected))
        .WandMaxWieldReq = Val(chWandMaxWieldReq.Text(chWandMaxWieldReq.Selected))
        .MaxWork = Val(chWandMaxWork.Text(chWandMaxWork.Selected))
    End With
    
    If g_Data.LootFilters.AddWandFilter(oFilter) Then
        'filterId updated
        Call AddWandFilterToList(oFilter)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".cmdAddWand_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub AddWandFilterToList(oFilter As clsLootFilter)
On Error GoTo ErrorHandler
    
    If Valid(oFilter) Then
        Dim i As Long
        i = lstLootWands.AddRow
        Call LinkRowToFilter(i, oFilter, FILTER_WAND)
        With lstLootWands
            .Data(COL_WAND_TYPE, i) = "Wand"
            .Data(COL_WAND_MANAC, i) = oFilter.WandManaC
            .Data(COL_WAND_MELEED, i) = oFilter.WandMeleeD
            .Data(COL_WAND_MAGICD, i) = oFilter.WandMagicD
            .Data(COL_WAND_PVM, i) = oFilter.WandPVM
            .Data(COL_WAND_WIELD_REQ, i) = oFilter.WandMaxWieldReq
            .Data(COL_WAND_MAX_WORK, i) = oFilter.MaxWork
            '.Data(COL_WAND_MAX_VALUE, i) = oFilter.MaxValue
            .ScrollPosition = i
        End With
        If oFilter.FilterEnabled Then
            lstLootWands.Data(COL_WAND_ENABLE, i) = 1
        Else
            lstLootWands.Data(COL_WAND_ENABLE, i) = 0
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".AddWandFilterToList - " & Err.Description
    Resume Fin
End Sub






'--------------------------------------- Exceptions List ------------------------------------
Public Sub UpdateExceptionList()
On Error GoTo ErrorHandler

    Dim objItem As acObject
    Dim i As Long
    
    Call lstExceptions.Clear
    For Each objItem In g_Data.Exceptions.Items
    
        i = lstExceptions.AddRow
        lstExceptions.Data(EX_COL_ICON, i, 1) = objItem.Icon + &H6000000
        lstExceptions.Data(EX_COL_NAME, i, 0) = objItem.Name
        lstExceptions.Data(EX_COL_GUID, i, 0) = objItem.Guid
        
        If Not g_Objects.Exists(objItem.Guid) Then
            lstExceptions.Color(EX_COL_NAME, i) = vbRed
        End If
        
    Next objItem

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".UpdateExceptionList - " & Err.Description
    Resume Fin
End Sub

Private Sub lstExceptions_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
On Error GoTo ErrorHandler
    
    Dim lGUID As Long
    lGUID = CLng(Val(lstExceptions.Data(EX_COL_GUID, nY)))
        
    If nX = EX_COL_ICON Then
        Call g_Hooks.SelectItem(lGUID)
    ElseIf nX = EX_COL_NAME Then
        If g_Data.Exceptions.RemoveItem(lGUID) Then
            Call UpdateExceptionList
        Else
            MyDebug "Unable to remove this item - GUID " & lGUID & " unknown in exception list"
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".lstExceptions_Change - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdAddException_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    If g_Hooks.CurrentSelection = 0 Then
        PrintErrorMessage "Please select an item first."
    Else
        Dim objItem As acObject
        
        If g_Objects.Items.Inv.Exists(g_Hooks.CurrentSelection, objItem) Then
            If g_Data.Exceptions.AddItem(objItem) Then
                Call UpdateExceptionList
            Else
                PrintErrorMessage "Failed to add item to exception list"
            End If
        Else
            PrintErrorMessage "This item doesnt belong to your inventory."
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage INTERFACE_NAME & ".lstExceptions_Change - " & Err.Description
    Resume Fin
End Sub

