VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const AUTO_SELECT = "Auto-Select"
Private Const TAG_ITEM_NAME = "item"
Private Const TAG_ENABLED = "enable"

Private Enum eRegenListColumns
    COL_ENABLED
    COL_ITEM_NAME
    COL_ICON_UP
    COL_ICON_DOWN
End Enum

Private InterfaceName As String

'Vitals
Public chkUseHealSpell As DecalControls.CheckBox
Public chkUseStamSpell As DecalControls.CheckBox
Public chkUseHealingKits As DecalControls.CheckBox
Public chkUseEmergItem As DecalControls.CheckBox
Public chkUseStamPotion As DecalControls.CheckBox
Public chkHealFellow As DecalControls.CheckBox
Public chkStamFellow As DecalControls.CheckBox
Public chkUsePeaceHeal As DecalControls.CheckBox

Public WithEvents sldMana As DecalControls.Slider
Attribute sldMana.VB_VarHelpID = -1
Public WithEvents sldStam As DecalControls.Slider
Attribute sldStam.VB_VarHelpID = -1
Public WithEvents sldHealth As DecalControls.Slider
Attribute sldHealth.VB_VarHelpID = -1
Public WithEvents sldCritHealth As DecalControls.Slider
Attribute sldCritHealth.VB_VarHelpID = -1
Public WithEvents sldFellowHealth As DecalControls.Slider
Attribute sldFellowHealth.VB_VarHelpID = -1
Public WithEvents sldFellowStam As DecalControls.Slider
Attribute sldFellowStam.VB_VarHelpID = -1

Public WithEvents txtAssistRange As DecalControls.Edit
Attribute txtAssistRange.VB_VarHelpID = -1

' Heal items tab
Public WithEvents lstHealItems As DecalControls.list
Attribute lstHealItems.VB_VarHelpID = -1
Public WithEvents btnAddHealItem As DecalControls.PushButton
Attribute btnAddHealItem.VB_VarHelpID = -1

' Stam items tab
Public WithEvents lstStamItems As DecalControls.list
Attribute lstStamItems.VB_VarHelpID = -1
Public WithEvents btnAddStamItem As DecalControls.PushButton
Attribute btnAddStamItem.VB_VarHelpID = -1

' Emergency items tab
Public WithEvents lstEmergItems As DecalControls.list
Attribute lstEmergItems.VB_VarHelpID = -1
Public WithEvents btnAddEmergItem As DecalControls.PushButton
Attribute btnAddEmergItem.VB_VarHelpID = -1

Public WithEvents cmdSetWand As DecalControls.PushButton
Attribute cmdSetWand.VB_VarHelpID = -1
Public lblWand As DecalControls.StaticText

Public chkLogoutOnDeath As DecalControls.CheckBox
Public chkLogoutOnTapers As DecalControls.CheckBox
Public chkLogoutOnKits As DecalControls.CheckBox
Public chkLogoutOnScarabs As DecalControls.CheckBox
Public chkLogoutOnArrows As DecalControls.CheckBox

'Should we cast a spell first?
Public chkLogoutCastSpell As DecalControls.CheckBox
Public chLogoutSpell As DecalControls.Choice
Attribute chLogoutSpell.VB_VarHelpID = -1

'Mage tab
Public WithEvents txtMageRadius As DecalControls.Edit
Attribute txtMageRadius.VB_VarHelpID = -1
Public WithEvents txtVulnRange As DecalControls.Edit
Attribute txtVulnRange.VB_VarHelpID = -1
Public WithEvents txtArcRange As DecalControls.Edit
Attribute txtArcRange.VB_VarHelpID = -1
Public WithEvents txtRingRange As DecalControls.Edit
Attribute txtRingRange.VB_VarHelpID = -1
Public WithEvents txtRingNum As DecalControls.Edit
Attribute txtRingNum.VB_VarHelpID = -1

Public chkVuln As DecalControls.CheckBox
Public chkImperil As DecalControls.CheckBox
Public chkYield As DecalControls.CheckBox
Public chkUseArcs As DecalControls.CheckBox
Public chkUseRings As DecalControls.CheckBox
Public chkUseVolleys As DecalControls.CheckBox
Public chkDebuffFirst As DecalControls.CheckBox
Public chkNoWar As DecalControls.CheckBox
Public WithEvents chWarLevel As DecalControls.Choice
Attribute chWarLevel.VB_VarHelpID = -1
Public WithEvents chVulnLevel As DecalControls.Choice
Attribute chVulnLevel.VB_VarHelpID = -1
Public WithEvents chDefaultDmg As DecalControls.Choice
Attribute chDefaultDmg.VB_VarHelpID = -1
Public WithEvents chkForceDefaultDmg As DecalControls.CheckBox
Attribute chkForceDefaultDmg.VB_VarHelpID = -1
Public chkEnableBlacklist As DecalControls.CheckBox

'Archer tab
Public WithEvents txtArcherRadius As DecalControls.Edit
Attribute txtArcherRadius.VB_VarHelpID = -1
Public WithEvents cmdSetBow As DecalControls.PushButton
Attribute cmdSetBow.VB_VarHelpID = -1
Public WithEvents cmdSetArrow As DecalControls.PushButton
Attribute cmdSetArrow.VB_VarHelpID = -1
Public WithEvents cmdSetHead As DecalControls.PushButton
Attribute cmdSetHead.VB_VarHelpID = -1
Public WithEvents cmdSetShaft As DecalControls.PushButton
Attribute cmdSetShaft.VB_VarHelpID = -1
Public lblBow As DecalControls.StaticText
Public lblArrows As DecalControls.StaticText
Public lblArrowHead As DecalControls.StaticText
Public lblArrowShaft As DecalControls.StaticText
Public chkEnableFletching As DecalControls.CheckBox
Public chkEnableArcherBlacklist As DecalControls.CheckBox

'Archer Arrows tab
Public WithEvents cmdArrowBludge As DecalControls.PushButton
Attribute cmdArrowBludge.VB_VarHelpID = -1
Public lblArrowBludge As DecalControls.StaticText
Public WithEvents cmdHeadBludge As DecalControls.PushButton
Attribute cmdHeadBludge.VB_VarHelpID = -1
Public lblHeadBludge As DecalControls.StaticText
Public WithEvents cmdArrowSlash As DecalControls.PushButton
Attribute cmdArrowSlash.VB_VarHelpID = -1
Public lblArrowSlash As DecalControls.StaticText
Public WithEvents cmdHeadSlash As DecalControls.PushButton
Attribute cmdHeadSlash.VB_VarHelpID = -1
Public lblHeadSlash As DecalControls.StaticText
Public WithEvents cmdArrowPierce As DecalControls.PushButton
Attribute cmdArrowPierce.VB_VarHelpID = -1
Public lblArrowPierce As DecalControls.StaticText
Public WithEvents cmdHeadPierce As DecalControls.PushButton
Attribute cmdHeadPierce.VB_VarHelpID = -1
Public lblHeadPierce As DecalControls.StaticText
Public WithEvents cmdArrowFire As DecalControls.PushButton
Attribute cmdArrowFire.VB_VarHelpID = -1
Public lblArrowFire As DecalControls.StaticText
Public WithEvents cmdHeadFire As DecalControls.PushButton
Attribute cmdHeadFire.VB_VarHelpID = -1
Public lblHeadFire As DecalControls.StaticText
Public WithEvents cmdArrowCold As DecalControls.PushButton
Attribute cmdArrowCold.VB_VarHelpID = -1
Public lblArrowCold As DecalControls.StaticText
Public WithEvents cmdHeadCold As DecalControls.PushButton
Attribute cmdHeadCold.VB_VarHelpID = -1
Public lblHeadCold As DecalControls.StaticText
Public WithEvents cmdArrowAcid As DecalControls.PushButton
Attribute cmdArrowAcid.VB_VarHelpID = -1
Public lblArrowAcid As DecalControls.StaticText
Public WithEvents cmdHeadAcid As DecalControls.PushButton
Attribute cmdHeadAcid.VB_VarHelpID = -1
Public lblHeadAcid As DecalControls.StaticText
Public WithEvents cmdArrowLight As DecalControls.PushButton
Attribute cmdArrowLight.VB_VarHelpID = -1
Public lblArrowLight As DecalControls.StaticText
Public WithEvents cmdHeadLight As DecalControls.PushButton
Attribute cmdHeadLight.VB_VarHelpID = -1
Public lblHeadLight As DecalControls.StaticText


'Melee tab
Public WithEvents txtMeleeRadius As DecalControls.Edit
Attribute txtMeleeRadius.VB_VarHelpID = -1
Public WithEvents cmdSetWeapon As DecalControls.PushButton
Attribute cmdSetWeapon.VB_VarHelpID = -1
Public WithEvents cmdSetShield As DecalControls.PushButton
Attribute cmdSetShield.VB_VarHelpID = -1
Public lblWeapon As DecalControls.StaticText
Public lblShield As DecalControls.StaticText

'Nav
Public WithEvents chkEnableNav As DecalControls.CheckBox
Attribute chkEnableNav.VB_VarHelpID = -1
Public WithEvents cmdShowNavEditor As DecalControls.PushButton
Attribute cmdShowNavEditor.VB_VarHelpID = -1
Public WithEvents chkEnableSticky As DecalControls.CheckBox
Attribute chkEnableSticky.VB_VarHelpID = -1
Public WithEvents cmdSetSticky As DecalControls.PushButton
Attribute cmdSetSticky.VB_VarHelpID = -1
Public WithEvents chkEnableFollow As DecalControls.CheckBox
Attribute chkEnableFollow.VB_VarHelpID = -1
Public WithEvents cmdSetFollow As DecalControls.PushButton
Attribute cmdSetFollow.VB_VarHelpID = -1
Public WithEvents chkNoCombatFollow As DecalControls.CheckBox
Attribute chkNoCombatFollow.VB_VarHelpID = -1

'Targetting
Public chkAttackAny As DecalControls.CheckBox
Public chkAttackSelected As DecalControls.CheckBox
Public chkAttackGroundOnly As DecalControls.CheckBox
Public chkSwitchIfBetterTarget As DecalControls.CheckBox
Public chkAttackVulnedOnly As DecalControls.CheckBox
Public chkAttackPK As DecalControls.CheckBox
Public chkDangerZone As DecalControls.CheckBox
Public txtDangerZone As DecalControls.Edit

'Elemental Weapons (wand, melee or bow)
Public chkAutoSwitch As DecalControls.CheckBox
Public WithEvents cmdWeapBludge As DecalControls.PushButton
Attribute cmdWeapBludge.VB_VarHelpID = -1
Public chkWeapBludge As DecalControls.CheckBox
Public lblWeapBludge As DecalControls.StaticText
Public WithEvents cmdWeapSlash As DecalControls.PushButton
Attribute cmdWeapSlash.VB_VarHelpID = -1
Public chkWeapSlash As DecalControls.CheckBox
Public lblWeapSlash As DecalControls.StaticText
Public WithEvents cmdWeapPierce As DecalControls.PushButton
Attribute cmdWeapPierce.VB_VarHelpID = -1
Public chkWeapPierce As DecalControls.CheckBox
Public lblWeapPierce As DecalControls.StaticText
Public WithEvents cmdWeapFire As DecalControls.PushButton
Attribute cmdWeapFire.VB_VarHelpID = -1
Public chkWeapFire As DecalControls.CheckBox
Public lblWeapFire As DecalControls.StaticText
Public WithEvents cmdWeapCold As DecalControls.PushButton
Attribute cmdWeapCold.VB_VarHelpID = -1
Public chkWeapCold As DecalControls.CheckBox
Public lblWeapCold As DecalControls.StaticText
Public WithEvents cmdWeapAcid As DecalControls.PushButton
Attribute cmdWeapAcid.VB_VarHelpID = -1
Public chkWeapAcid As DecalControls.CheckBox
Public lblWeapAcid As DecalControls.StaticText
Public WithEvents cmdWeapLight As DecalControls.PushButton
Attribute cmdWeapLight.VB_VarHelpID = -1
Public chkWeapLight As DecalControls.CheckBox
Public lblWeapLight As DecalControls.StaticText


'Fellow
Public chkEnableAutoFellow As DecalControls.CheckBox
Public txtFellowPassword As DecalControls.Edit
Public chkAutoFellowList As DecalControls.CheckBox
Public WithEvents FellowList As DecalControls.list
Attribute FellowList.VB_VarHelpID = -1

'Spell levels
Public chMacroSpellsLevel As DecalControls.Choice 'revit spells
Public chHealSpellsLevel As DecalControls.Choice
Public chEmergencySpellsLevel As DecalControls.Choice


'Options
Public chkEnableTurbo As DecalControls.CheckBox
Public chkEnableDOT As DecalControls.CheckBox
'Public WithEvents sldMacroTimer As DecalControls.Slider

'Misc
Public WithEvents chAttackHeight As DecalControls.Choice
Attribute chAttackHeight.VB_VarHelpID = -1
Public chkUseManaCharge As DecalControls.CheckBox
Public chkUseManaStone As DecalControls.CheckBox
Public chkUseMassiveCharges As DecalControls.CheckBox
Public chkRechargeManaStones As DecalControls.CheckBox
'Public headerButton As DecalControls.PushButton
Public WithEvents btnStopSounds As DecalControls.PushButton
Attribute btnStopSounds.VB_VarHelpID = -1
Public WithEvents btnResetHUD As DecalControls.PushButton
Attribute btnResetHUD.VB_VarHelpID = -1
Public WithEvents btnExportInv As DecalControls.PushButton
Attribute btnExportInv.VB_VarHelpID = -1


Public Function Init(Optional ProfileName As String = "Default") As Boolean
On Error GoTo ErrorHandler

    'Set interface module name
    InterfaceName = "UIMacro"
    
    MyDebug InterfaceName & ".Init() -- Begin"
    
    'Decal Controls initialisation
    
    Set chkAttackAny = g_MainView.Control("chkAttackAny")
    Set chkAttackGroundOnly = g_MainView.Control("chkAttackGroundOnly")
    Set chkAttackSelected = g_MainView.Control("chkAttackSelected")
    Set chkSwitchIfBetterTarget = g_MainView.Control("chkSwitchIfBetterTarget")
    Set chkAttackVulnedOnly = g_MainView.Control("chkAttackVulnedOnly")
    Set chkAttackPK = g_MainView.Control("chkAttackPK")
    Set chkDangerZone = g_MainView.Control("chkDangerZone")
    Set txtDangerZone = g_MainView.Control("txtDangerZone")
     
    Set chkUseHealSpell = g_MainView.Control("chkUseHealSpell")
    Set chkUseStamSpell = g_MainView.Control("chkUseStamSpell")
    Set chkUseHealingKits = g_MainView.Control("chkUseHealingKits")
    Set chkUseEmergItem = g_MainView.Control("chkUseEmergItem")
    Set chkUseStamPotion = g_MainView.Control("chkUseStamPotion")
    Set chkUsePeaceHeal = g_MainView.Control("chkUsePeaceHeal")
    
    Set sldMana = g_MainView.Control("sldMana")
    Set sldStam = g_MainView.Control("sldStam")
    Set sldHealth = g_MainView.Control("sldHealth")
    Set sldCritHealth = g_MainView.Control("sldCritHealth")
    
    Set chkHealFellow = g_MainView.Control("chkHealFellow")
    Set chkStamFellow = g_MainView.Control("chkStamFellow")
    Set sldFellowHealth = g_MainView.Control("sldFellowHealth")
    Set sldFellowStam = g_MainView.Control("sldFellowStam")
    
    Set lstHealItems = g_MainView.Control("lstHealItems")
    Set btnAddHealItem = g_MainView.Control("btnAddHealItem")
    
    Set lstStamItems = g_MainView.Control("lstStamItems")
    Set btnAddStamItem = g_MainView.Control("btnAddStamItem")
    
    Set lstEmergItems = g_MainView.Control("lstEmergItems")
    Set btnAddEmergItem = g_MainView.Control("btnAddEmergItem")
    
    Set chkLogoutOnDeath = g_MainView.Control("chkLogoutOnDeath")
    Set cmdSetWand = g_MainView.Control("cmdSetWand")
    Set lblWand = g_MainView.Control("lblWand")
    
    'Mage tab
    Set txtMageRadius = g_MainView.Control("txtMageRadius")
    Set txtVulnRange = g_MainView.Control("txtVulnRange")
    Set txtArcRange = g_MainView.Control("txtArcRange")
    Set txtRingRange = g_MainView.Control("txtRingRange")
    Set txtRingNum = g_MainView.Control("txtRingNum")
    Set txtAssistRange = g_MainView.Control("txtAssistRange")
    Set chkVuln = g_MainView.Control("chkVuln")
    Set chkImperil = g_MainView.Control("chkImperil")
    Set chkYield = g_MainView.Control("chkYield")
    Set chkUseArcs = g_MainView.Control("chkUseArcs")
    Set chkUseRings = g_MainView.Control("chkUseRings")
    Set chkUseVolleys = g_MainView.Control("chkUseVolleys")
    Set chkDebuffFirst = g_MainView.Control("chkDebuffFirst")
    Set chkNoWar = g_MainView.Control("chkNoWar")
    Set chWarLevel = g_MainView.Control("chWarLevel")
    Set chVulnLevel = g_MainView.Control("chVulnLevel")
    Set chDefaultDmg = g_MainView.Control("chDefaultDmg")
    Set chkEnableBlacklist = g_MainView.Control("chkEnableBlacklist")
    Set chkForceDefaultDmg = g_MainView.Control("chkForceDefaultDmg")
    Set chAttackHeight = g_MainView.Control("chAttackHeight")

    'Archer tab
    Set txtArcherRadius = g_MainView.Control("txtArcherRadius")
    Set cmdSetBow = g_MainView.Control("cmdSetBow")
    Set cmdSetArrow = g_MainView.Control("cmdSetArrow")
    Set cmdSetHead = g_MainView.Control("cmdSetHead")
    Set cmdSetShaft = g_MainView.Control("cmdSetShaft")
    Set lblBow = g_MainView.Control("lblBow")
    Set lblArrows = g_MainView.Control("lblArrows")
    Set lblArrowHead = g_MainView.Control("lblArrowHead")
    Set lblArrowShaft = g_MainView.Control("lblArrowShaft")
    Set chkEnableFletching = g_MainView.Control("chkEnableFletching")
    Set chkEnableArcherBlacklist = g_MainView.Control("chkEnableArcherBlacklist")
    ' Arrows
    Set cmdArrowBludge = g_MainView.Control("cmdArrowBludge")
    Set lblArrowBludge = g_MainView.Control("lblArrowBludge")
    Set cmdHeadBludge = g_MainView.Control("cmdHeadBludge")
    Set lblHeadBludge = g_MainView.Control("lblHeadBludge")
    
    Set cmdArrowSlash = g_MainView.Control("cmdArrowSlash")
    Set lblArrowSlash = g_MainView.Control("lblArrowSlash")
    Set cmdHeadSlash = g_MainView.Control("cmdHeadSlash")
    Set lblHeadSlash = g_MainView.Control("lblHeadSlash")
    
    Set cmdArrowPierce = g_MainView.Control("cmdArrowPierce")
    Set lblArrowPierce = g_MainView.Control("lblArrowPierce")
    Set cmdHeadPierce = g_MainView.Control("cmdHeadPierce")
    Set lblHeadPierce = g_MainView.Control("lblHeadPierce")
    
    Set cmdArrowFire = g_MainView.Control("cmdArrowFire")
    Set lblArrowFire = g_MainView.Control("lblArrowFire")
    Set cmdHeadFire = g_MainView.Control("cmdHeadFire")
    Set lblHeadFire = g_MainView.Control("lblHeadFire")
    
    Set cmdArrowCold = g_MainView.Control("cmdArrowCold")
    Set lblArrowCold = g_MainView.Control("lblArrowCold")
    Set cmdHeadCold = g_MainView.Control("cmdHeadCold")
    Set lblHeadCold = g_MainView.Control("lblHeadCold")
    
    Set cmdArrowAcid = g_MainView.Control("cmdArrowAcid")
    Set lblArrowAcid = g_MainView.Control("lblArrowAcid")
    Set cmdHeadAcid = g_MainView.Control("cmdHeadAcid")
    Set lblHeadAcid = g_MainView.Control("lblHeadAcid")
    
    Set cmdArrowLight = g_MainView.Control("cmdArrowLight")
    Set lblArrowLight = g_MainView.Control("lblArrowLight")
    Set cmdHeadLight = g_MainView.Control("cmdHeadLight")
    Set lblHeadLight = g_MainView.Control("lblHeadLight")
    
    
    'Melee tab
    Set txtMeleeRadius = g_MainView.Control("txtMeleeRadius")
    Set cmdSetWeapon = g_MainView.Control("cmdSetWeapon")
    Set cmdSetShield = g_MainView.Control("cmdSetShield")
    Set lblWeapon = g_MainView.Control("lblWeapon")
    Set lblShield = g_MainView.Control("lblShield")
    
    'Items Tab
    Set chkAutoSwitch = g_MainView.Control("chkAutoSwitch")
    
    Set cmdWeapBludge = g_MainView.Control("cmdWeapBludge")
    Set chkWeapBludge = g_MainView.Control("chkWeapBludge")
    Set lblWeapBludge = g_MainView.Control("lblWeapBludge")
    Set cmdWeapSlash = g_MainView.Control("cmdWeapSlash")
    Set chkWeapSlash = g_MainView.Control("chkWeapSlash")
    Set lblWeapSlash = g_MainView.Control("lblWeapSlash")
    Set cmdWeapPierce = g_MainView.Control("cmdWeapPierce")
    Set chkWeapPierce = g_MainView.Control("chkWeapPierce")
    Set lblWeapPierce = g_MainView.Control("lblWeapPierce")
    Set cmdWeapFire = g_MainView.Control("cmdWeapFire")
    Set chkWeapFire = g_MainView.Control("chkWeapFire")
    Set lblWeapFire = g_MainView.Control("lblWeapFire")
    Set cmdWeapCold = g_MainView.Control("cmdWeapCold")
    Set chkWeapCold = g_MainView.Control("chkWeapCold")
    Set lblWeapCold = g_MainView.Control("lblWeapCold")
    Set cmdWeapAcid = g_MainView.Control("cmdWeapAcid")
    Set chkWeapAcid = g_MainView.Control("chkWeapAcid")
    Set lblWeapAcid = g_MainView.Control("lblWeapAcid")
    Set cmdWeapLight = g_MainView.Control("cmdWeapLight")
    Set chkWeapLight = g_MainView.Control("chkWeapLight")
    Set lblWeapLight = g_MainView.Control("lblWeapLight")

    Set chkLogoutOnTapers = g_MainView.Control("chkLogoutOnTapers")
    Set chkLogoutOnKits = g_MainView.Control("chkLogoutOnKits")
    Set chkLogoutOnScarabs = g_MainView.Control("chkLogoutOnScarabs")
    Set chkLogoutOnArrows = g_MainView.Control("chkLogoutOnArrows")
    
    Set chkLogoutCastSpell = g_MainView.Control("chkLogoutCastSpell")
    Set chLogoutSpell = g_MainView.Control("chLogoutSpell")

    'Nav tab
    Set chkEnableNav = g_MainView.Control("chkEnableNav")
    Set cmdShowNavEditor = g_MainView.Control("cmdShowNavEditor")
    Set chkEnableSticky = g_MainView.Control("chkEnableSticky")
    Set cmdSetSticky = g_MainView.Control("cmdSetSticky")
    Set chkEnableFollow = g_MainView.Control("chkEnableFollow")
    Set cmdSetFollow = g_MainView.Control("cmdSetFollow")
    Set chkNoCombatFollow = g_MainView.Control("chkNoCombatFollow")
    
    'Options
    Set chMacroSpellsLevel = g_MainView.Control("chMacroSpellsLevel")
    Set chHealSpellsLevel = g_MainView.Control("chHealSpellsLevel")
    Set chEmergencySpellsLevel = g_MainView.Control("chEmergencySpellsLevel")

    Set chkEnableTurbo = g_MainView.Control("chkEnableTurbo")
    Set chkEnableDOT = g_MainView.Control("chkEnableDOT")
    'Set sldMacroTimer = g_MainView.Control("sldMacroTimer")
    
    'Fellow tab
    Set chkEnableAutoFellow = g_MainView.Control("chkEnableAutoFellow")
    Set txtFellowPassword = g_MainView.Control("txtFellowPassword")
    Set chkAutoFellowList = g_MainView.Control("chkAutoFellowList")
    Set FellowList = g_MainView.Control("FellowList")

    'Misc
    Set chkUseManaCharge = g_MainView.Control("chkUseManaCharge")
    Set chkUseManaStone = g_MainView.Control("chkUseManaStone")
    Set chkUseMassiveCharges = g_MainView.Control("chkUseMassiveCharges")
    Set chkRechargeManaStones = g_MainView.Control("chkRechargeManaStones")
    
    Set btnStopSounds = g_MainView.Control("btnStopSounds")
    Set btnResetHUD = g_MainView.Control("btnResetHUD")
    Set btnExportInv = g_MainView.Control("btnExportInv")
    
    'Set headerButton = g_MainView.Control("headerButton")
    
    'Init lists
    Dim i As Integer
    Call chDefaultDmg.Clear
    For i = DMG_SLASHING To DMG_LIGHTNING
        chDefaultDmg.AddChoice GetDamageString(i)
    Next i
    chDefaultDmg.AddChoice AUTO_SELECT
    
    Init = True
    MyDebug InterfaceName & ".Init() -- End"
    
Fin:
    Exit Function
ErrorHandler:
    Init = False
    PrintErrorMessage InterfaceName & ".Init - " & Err.Description
    Resume Fin
End Function

Public Function LoadControlsValue(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler

    LoadControlsValue = False
    
    '***************************************************************************
    
    'Main settings
    With g_Settings.Profile.MacroCfg
    
        chkAttackAny.Checked = .GetValue("chkAttackAny", False)
        chkAttackGroundOnly.Checked = .GetValue("chkAttackGroundOnly", False)
        chkAttackSelected.Checked = .GetValue("chkAttackSelected", False)
        chkSwitchIfBetterTarget.Checked = .GetValue("chkSwitchIfBetterTarget", True)
        chkAttackVulnedOnly.Checked = .GetValue("chkAttackVulnedOnly", True)  'vulneds first
        chkAttackPK.Checked = .GetValue("chkAttackPK", True)
        chkDangerZone.Checked = .GetValue("chkDangerZone", True)
        txtDangerZone.Text = .GetValue("txtDangerZone", 5)
        
        sldMana.SliderPosition = .GetValue("sldMana", 50)
        sldStam.SliderPosition = .GetValue("sldStam", 50)
        sldHealth.SliderPosition = .GetValue("sldHealth", 50)
        sldCritHealth.SliderPosition = .GetValue("sldCritHealth", 50)
        sldFellowHealth.SliderPosition = .GetValue("sldFellowHealth", 30)
        sldFellowStam.SliderPosition = .GetValue("sldFellowStam", 10)
        
        'Mage
        txtMageRadius.Text = .GetValue("txtMageRadius", 30)
        txtVulnRange.Text = .GetValue("txtVulnRange", 30)
        txtArcRange.Text = .GetValue("txtArcRange", 5)
        txtRingRange.Text = .GetValue("txtRingRange", 5)
        txtAssistRange.Text = .GetValue("txtAssistRange", 20)
        
        txtRingNum.Text = .GetValue("txtRingNum", 4)
    
        chkVuln.Checked = .GetValue("chkVuln", False)
        chkImperil.Checked = .GetValue("chkImperil", False)
        chkYield.Checked = .GetValue("chkYield", False)
        chkUseArcs.Checked = .GetValue("chkUseArcs", False)
        chkUseRings.Checked = .GetValue("chkUseRings", False)
        chkUseVolleys.Checked = .GetValue("chkUseVolleys", False)
        chkDebuffFirst.Checked = .GetValue("chkDebuffFirst", False)
        chkNoWar.Checked = .GetValue("chkNoWar", False)
        chkForceDefaultDmg.Checked = .GetValue("chkForceDefaultDmg", False)
        chkEnableBlacklist.Checked = .GetValue("chkEnableBlacklist", True)
         
        chWarLevel.Selected = .GetValue("chWarLevel", 6)
        chVulnLevel.Selected = .GetValue("chVulnLevel", 6)
        chDefaultDmg.Selected = .GetValue("chDefaultDmg", 0)
        
        'Archer
        txtArcherRadius.Text = .GetValue("txtArcherRadius", 30)
        chkEnableFletching.Checked = .GetValue("chkEnableFletching", False)
        chkEnableArcherBlacklist.Checked = .GetValue("chkEnableArcherBlacklist", False)
    
        'Melee
        txtMeleeRadius.Text = .GetValue("txtMeleeRadius", 30)
        
        'Nav Tab
        chkEnableNav.Checked = .GetValue("chkEnableNav", False)
        ' Don't save/load sticky position
        'chkEnableSticky.Checked = .GetValue("chkEnableSticky", False)
        'chkEnableFollow.Checked = .GetValue("chkEnableFollow", False)
        chkNoCombatFollow.Checked = .GetValue("chkNoCombatFollow", False)
        
        'Load route
        g_Nav.Route.Description = .GetValue(ID_ROUTE_NAME, "Default Route")
        g_Nav.NavType = .GetValue(ID_ROUTE_NAVTYPE, NAVTYPE_LOOP)
        MyDebug "---- NavType : " & g_Nav.NavType
        If RouteExist(g_Nav.Route.Description) Then
            If g_Nav.Route.LoadRoute(GetRouteFilePath(g_Nav.Route.Description)) Then
                MyDebug "Loaded route : " & g_Nav.Route.Description & " (" & g_Nav.Route.NumWP & " waypoints)"
                'PrintMessage "Loaded route : " & g_Nav.Route.Description & " (" & g_Nav.Route.NumWP & " waypoints)"
            Else
                PrintErrorMessage "Failed to load route : " & g_Nav.Route.Description
            End If
        End If
        
        'Load Weapon/Wand/Bow guids
        g_Data.WandGUID = .GetValue(ID_WAND, 0)
        g_Data.WeaponGUID = .GetValue(ID_WEAPON, 0)
        g_Data.ShieldGUID = .GetValue(ID_SHIELD, 0)
        g_Data.BowGUID = .GetValue(ID_BOW, 0)
        g_Data.Arrows = .GetValue(ID_ARROWS, TAG_NOT_SET)
        g_Data.ArrowHead = .GetValue(ID_ARROWHEAD, TAG_NOT_SET)
        g_Data.ArrowShaft = .GetValue(ID_ARROWSHAFT, TAG_NOT_SET)
        
        g_Data.m_sArrowBludge = .GetValue("ArrowBludge", TAG_NOT_SET)
        g_Data.m_sHeadBludge = .GetValue("HeadBludge", TAG_NOT_SET)
        g_Data.m_sArrowSlash = .GetValue("ArrowSlash", TAG_NOT_SET)
        g_Data.m_sHeadSlash = .GetValue("HeadSlash", TAG_NOT_SET)
        g_Data.m_sArrowPierce = .GetValue("ArrowPierce", TAG_NOT_SET)
        g_Data.m_sHeadPierce = .GetValue("HeadPierce", TAG_NOT_SET)
        g_Data.m_sArrowFire = .GetValue("ArrowFire", TAG_NOT_SET)
        g_Data.m_sHeadFire = .GetValue("HeadFire", TAG_NOT_SET)
        g_Data.m_sArrowCold = .GetValue("ArrowCold", TAG_NOT_SET)
        g_Data.m_sHeadCold = .GetValue("HeadCold", TAG_NOT_SET)
        g_Data.m_sArrowAcid = .GetValue("ArrowAcid", TAG_NOT_SET)
        g_Data.m_sHeadAcid = .GetValue("HeadAcid", TAG_NOT_SET)
        g_Data.m_sArrowLight = .GetValue("ArrowLight", TAG_NOT_SET)
        g_Data.m_sHeadLight = .GetValue("HeadLight", TAG_NOT_SET)
        
        chkAutoSwitch.Checked = .GetValue("chkAutoSwitch", False)
        
        g_Data.SetWeapBludgeGUID = .GetValue("WeapBludge", 0)
        chkWeapBludge.Checked = .GetValue("chkWeapBludge", False)
        g_Data.SetWeapSlashGUID = .GetValue("WeapSlash", 0)
        chkWeapSlash.Checked = .GetValue("chkWeapSlash", False)
        g_Data.SetWeapPierceGUID = .GetValue("WeapPierce", 0)
        chkWeapPierce.Checked = .GetValue("chkWeapPierce", False)
        g_Data.SetWeapFireGUID = .GetValue("WeapFire", 0)
        chkWeapFire.Checked = .GetValue("chkWeapFire", False)
        g_Data.SetWeapColdGUID = .GetValue("WeapCold", 0)
        chkWeapCold.Checked = .GetValue("chkWeapCold", False)
        g_Data.SetWeapAcidGUID = .GetValue("WeapAcid", 0)
        chkWeapAcid.Checked = .GetValue("chkWeapAcid", False)
        g_Data.SetWeapLightGUID = .GetValue("WeapLight", 0)
        chkWeapLight.Checked = .GetValue("chkWeapLight", False)
        
        Call UpdateMacroItemsLabels
        
        'SPK
        chAttackHeight.Selected = .GetValue("chAttackHeight", 0)

    End With
        
    With g_Settings.Profile.Config
    
        chkLogoutOnTapers.Checked = g_Settings.GetValue("chkLogoutOnTapers", True)
        chkLogoutOnKits.Checked = g_Settings.GetValue("chkLogoutOnKits", False)
        chkLogoutOnScarabs.Checked = g_Settings.GetValue("chkLogoutOnScarabs", True)
        chkLogoutOnArrows.Checked = g_Settings.GetValue("chkLogoutOnArrows", False)
        
        chkLogoutCastSpell.Checked = g_Settings.GetValue("chkLogoutCastSpell", False)
        
        Dim objFam As clsSpellFamily
        
        Call chLogoutSpell.Clear
        For Each objFam In g_Spells.Items.Families
            If objFam.mType = SPELLTYPE_TRANSPORT Then
                Call chLogoutSpell.AddChoice(objFam.mFamily)
            End If
        Next objFam
        If chLogoutSpell.ChoiceCount > 0 Then
            chLogoutSpell.Selected = 0
        End If

        chkUseHealSpell.Checked = g_Settings.GetValue("chkUseHealSpell", True)
        chkUseStamSpell.Checked = g_Settings.GetValue("chkUseStamSpell", True)
        chkUseHealingKits.Checked = g_Settings.GetValue("chkUseHealingKits", True)
        chkUseEmergItem.Checked = g_Settings.GetValue("chkUseEmergItem", True)
        chkUseStamPotion.Checked = g_Settings.GetValue("chkUseStamPotion", False)
        chkUsePeaceHeal.Checked = g_Settings.GetValue("chkUsePeaceHeal", True)
        'chkUsePeaceHeal.Unchecked = g_Settings.GetValue("chkUsePeaceHeal", False)
        chkHealFellow.Checked = g_Settings.GetValue("chkHealFellow", False)
        chkStamFellow.Checked = g_Settings.GetValue("chkStamFellow", False)
    
        chkLogoutOnDeath.Checked = g_Settings.GetValue("chkLogoutOnDeath", True)
        chkEnableTurbo.Checked = g_Settings.GetValue("chkEnableTurbo", True)
        chkEnableDOT.Checked = g_Settings.GetValue("chkEnableDOT", False)

        'Fellow tab
        chkEnableAutoFellow.Checked = g_Settings.GetValue("chkEnableAutoFellow", True)
        txtFellowPassword.Text = g_Settings.GetValue("txtFellowPassword", "xp")
        chkAutoFellowList.Checked = g_Settings.GetValue("chkAutoFellowList", False)

        'Options
        chMacroSpellsLevel.Selected = g_Settings.GetValue("chMacroSpellsLevel", 6)
        chHealSpellsLevel.Selected = g_Settings.GetValue("chHealSpellsLevel", 6)
        chEmergencySpellsLevel.Selected = g_Settings.GetValue("chEmergencySpellsLevel", 2)
        
        'sldCastBreaker.SliderPosition = g_Settings.GetValue("sldCastBreaker", 2900)
        'sldMacroTimer.SliderPosition = g_Settings.GetValue("sldMacroTimer", 200)
        
        chkUseManaCharge.Checked = g_Settings.GetValue("chkUseManaCharge", True)
        chkUseManaStone.Checked = g_Settings.GetValue("chkUseManaStone", True)
        chkUseMassiveCharges.Checked = g_Settings.GetValue("chkUseMassiveCharges", False)
        chkRechargeManaStones.Checked = g_Settings.GetValue("chkRechargeManaStones", False)
                
    End With
    
    Call LoadItemsList(lstHealItems, GetHealListPath)
    Call LoadItemsList(lstStamItems, GetStamListPath)
    Call LoadItemsList(lstEmergItems, GetEmergListPath)
    
    '***************************************************************************
    
    LoadControlsValue = True
    
Fin:
    Exit Function
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage "clsUIMain.LoadControlsValue - " & Err.Description
    Resume Fin
End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    
    MyDebug "[" & InterfaceName & "] Saving controls settings"
    
    With g_Settings.Profile.Config
    
        .SaveCheckbox chkUseHealSpell, "chkUseHealSpell"
        .SaveCheckbox chkUseStamSpell, "chkUseStamSpell"
        .SaveCheckbox chkUseHealingKits, "chkUseHealingKits"
        .SaveCheckbox chkUseEmergItem, "chkUseEmergItem"
        .SaveCheckbox chkUseStamPotion, "chkUseStamPotion"
        .SaveCheckbox chkUsePeaceHeal, "chkUsePeaceHeal"
        .SaveCheckbox chkHealFellow, "chkHealFellow"
        .SaveCheckbox chkStamFellow, "chkStamFellow"
        
        .SaveCheckbox chkLogoutOnDeath, "chkLogoutOnDeath"
        
        'Fellow
        .SaveCheckbox chkEnableAutoFellow, "chkEnableAutoFellow"
        .SaveTextbox txtFellowPassword, "txtFellowPassword"
        .SaveCheckbox chkAutoFellowList, "chkAutoFellowList"
        
        'Options
        .SaveCheckbox chkLogoutOnTapers, "chkLogoutOnTapers"
        .SaveCheckbox chkLogoutOnKits, "chkLogoutOnKits"
        .SaveCheckbox chkLogoutOnScarabs, "chkLogoutOnScarabs"
        .SaveCheckbox chkLogoutOnArrows, "chkLogoutOnArrows"
        
        .SaveCheckbox chkLogoutCastSpell, "chkLogoutCastSpell"
        
        'Spells lvl
        .SaveChoice chMacroSpellsLevel, "chMacroSpellsLevel"
        .SaveChoice chHealSpellsLevel, "chHealSpellsLevel"
        .SaveChoice chEmergencySpellsLevel, "chEmergencySpellsLevel"
        
        .SaveCheckbox chkEnableTurbo, "chkEnableTurbo"
        .SaveCheckbox chkEnableDOT, "chkEnableDOT"
        '.SaveSlider sldMacroTimer, "sldMacroTimer"
        
        .SaveCheckbox chkUseManaCharge, "chkUseManaCharge"
        .SaveCheckbox chkUseManaStone, "chkUseManaStone"
        .SaveCheckbox chkUseMassiveCharges, "chkUseMassiveCharges"
        .SaveCheckbox chkRechargeManaStones, "chkRechargeManaStones"
        
    End With
        
    With g_Settings.Profile.MacroCfg

        'Main settings
        .SaveCheckbox chkAttackAny, "chkAttackAny"
        .SaveCheckbox chkAttackGroundOnly, "chkAttackGroundOnly"
        .SaveCheckbox chkAttackSelected, "chkAttackSelected"
        .SaveCheckbox chkSwitchIfBetterTarget, "chkSwitchIfBetterTarget"
        .SaveCheckbox chkAttackVulnedOnly, "chkAttackVulnedOnly"
        .SaveCheckbox chkAttackPK, "chkAttackPK"
        .SaveCheckbox chkDangerZone, "chkDangerZone"
        .SaveTextbox txtDangerZone, "txtDangerZone"
        
        
        .SaveSlider sldMana, "sldMana"
        .SaveSlider sldStam, "sldStam"
        .SaveSlider sldHealth, "sldHealth"
        .SaveSlider sldCritHealth, "sldCritHealth"
        .SaveSlider sldFellowHealth, "sldFellowHealth"
        .SaveSlider sldFellowStam, "sldFellowStam"
        
        'save wand guid
        .SaveValue ID_WAND, g_Data.WandGUID
        
        'Mage
        .SaveTextbox txtMageRadius, "txtMageRadius"
        .SaveTextbox txtVulnRange, "txtVulnRange"
        .SaveTextbox txtArcRange, "txtArcRange"
        .SaveTextbox txtRingRange, "txtRingRange"
        .SaveTextbox txtAssistRange, "txtAssistRange"
        
        .SaveTextbox txtRingNum, "txtRingNum"

        .SaveCheckbox chkVuln, "chkVuln"
        .SaveCheckbox chkImperil, "chkImperil"
        .SaveCheckbox chkYield, "chkYield"
        .SaveCheckbox chkUseArcs, "chkUseArcs"
        .SaveCheckbox chkUseRings, "chkUseRings"
        .SaveCheckbox chkUseVolleys, "chkUseVolleys"
        .SaveCheckbox chkDebuffFirst, "chkDebuffFirst"
        .SaveCheckbox chkNoWar, "chkNoWar"
        .SaveCheckbox chkForceDefaultDmg, "chkForceDefaultDmg"
        .SaveCheckbox chkEnableBlacklist, "chkEnableBlacklist"
        
        .SaveChoice chWarLevel, "chWarLevel"
        .SaveChoice chVulnLevel, "chVulnLevel"
        .SaveChoice chDefaultDmg, "chDefaultDmg"
        
        'Archer
        .SaveTextbox txtArcherRadius, "txtArcherRadius"
        .SaveCheckbox chkEnableFletching, "chkEnableFletching"
        .SaveCheckbox chkEnableArcherBlacklist, "chkEnableArcherBlacklist"
        
        'save bow guid
        .SaveValue ID_BOW, g_Data.BowGUID
        
        'Melee
        .SaveTextbox txtMeleeRadius, "txtMeleeRadius"
        
        'save weapon/shield guid
        .SaveValue ID_WEAPON, g_Data.WeaponGUID
        .SaveValue ID_SHIELD, g_Data.ShieldGUID
        
        'fletching stuff
        .SaveValue ID_ARROWS, g_Data.Arrows
        .SaveValue ID_ARROWHEAD, g_Data.ArrowHead
        .SaveValue ID_ARROWSHAFT, g_Data.ArrowShaft
        
        .SaveValue "ArrowBludge", g_Data.m_sArrowBludge
        .SaveValue "HeadBludge", g_Data.m_sHeadBludge
        .SaveValue "ArrowSlash", g_Data.m_sArrowSlash
        .SaveValue "HeadSlash", g_Data.m_sHeadSlash
        .SaveValue "ArrowPierce", g_Data.m_sArrowPierce
        .SaveValue "HeadPierce", g_Data.m_sHeadPierce
        .SaveValue "ArrowFire", g_Data.m_sArrowFire
        .SaveValue "HeadFire", g_Data.m_sHeadFire
        .SaveValue "ArrowCold", g_Data.m_sArrowCold
        .SaveValue "HeadCold", g_Data.m_sHeadCold
        .SaveValue "ArrowAcid", g_Data.m_sArrowAcid
        .SaveValue "HeadAcid", g_Data.m_sHeadAcid
        .SaveValue "ArrowLight", g_Data.m_sArrowLight
        .SaveValue "HeadLight", g_Data.m_sHeadLight
        
        'Items
        .SaveCheckbox chkAutoSwitch, "chkAutoSwitch"
        
        .SaveValue "WeapBludge", g_Data.WeapBludgeGUID
        .SaveCheckbox chkWeapBludge, "chkWeapBludge"
        .SaveValue "WeapSlash", g_Data.WeapSlashGUID
        .SaveCheckbox chkWeapSlash, "chkWeapSlash"
        .SaveValue "WeapPierce", g_Data.WeapPierceGUID
        .SaveCheckbox chkWeapPierce, "chkWeapPierce"
        .SaveValue "WeapFire", g_Data.WeapFireGUID
        .SaveCheckbox chkWeapFire, "chkWeapFire"
        .SaveValue "WeapCold", g_Data.WeapColdGUID
        .SaveCheckbox chkWeapCold, "chkWeapCold"
        .SaveValue "WeapAcid", g_Data.WeapAcidGUID
        .SaveCheckbox chkWeapAcid, "chkWeapAcid"
        .SaveValue "WeapLight", g_Data.WeapLightGUID
        .SaveCheckbox chkWeapLight, "chkWeapLight"
        
        'Nav
        .SaveCheckbox chkEnableNav, "chkEnableNav"
        .SaveValue ID_ROUTE_NAME, g_Nav.Route.Description   'save current route name
        .SaveValue ID_ROUTE_NAVTYPE, g_Nav.NavType          'save current route navtype
        .SaveCheckbox chkNoCombatFollow, "chkNoCombatFollow"
    
        .SaveChoice chAttackHeight, "chAttackHeight"
    End With
    
    Call SaveItemsList(lstHealItems, GetHealListPath)
    Call SaveItemsList(lstStamItems, GetStamListPath)
    Call SaveItemsList(lstEmergItems, GetEmergListPath)
    
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
    
    Set chkAttackAny = Nothing
    Set chkAttackGroundOnly = Nothing
    Set chkAttackSelected = Nothing
    Set chkSwitchIfBetterTarget = Nothing
    Set chkAttackVulnedOnly = Nothing
    Set chkAttackPK = Nothing
    Set chkDangerZone = Nothing
    Set txtDangerZone = Nothing
    
    Set chkUseHealSpell = Nothing
    Set chkUseStamSpell = Nothing
    Set chkUseHealingKits = Nothing
    Set chkUseEmergItem = Nothing
    Set chkUseStamPotion = Nothing
    Set chkUsePeaceHeal = Nothing
    Set chkHealFellow = Nothing
    Set chkStamFellow = Nothing
    
    Set sldMana = Nothing
    Set sldStam = Nothing
    Set sldHealth = Nothing
    Set sldCritHealth = Nothing
    Set sldFellowHealth = Nothing
    Set sldFellowStam = Nothing
    
    Set lstHealItems = Nothing
    Set lstStamItems = Nothing
    Set btnAddHealItem = Nothing
    Set btnAddStamItem = Nothing
 
    Set cmdSetWand = Nothing
    Set lblWand = Nothing
    
    Set chkLogoutOnDeath = Nothing
    
    'Mage tab
    Set txtMageRadius = Nothing
    Set txtVulnRange = Nothing
    Set txtArcRange = Nothing
    Set txtRingRange = Nothing
    Set txtAssistRange = Nothing
    Set txtRingNum = Nothing
    Set chkVuln = Nothing
    Set chkImperil = Nothing
    Set chkYield = Nothing
    Set chkUseArcs = Nothing
    Set chkUseRings = Nothing
    Set chkUseVolleys = Nothing
    Set chkDebuffFirst = Nothing
    Set chkNoWar = Nothing
    Set chWarLevel = Nothing
    Set chVulnLevel = Nothing
    Set chDefaultDmg = Nothing
    Set chkForceDefaultDmg = Nothing
    
    'Archer tab
    Set txtArcherRadius = Nothing
    Set cmdSetBow = Nothing
    Set cmdSetArrow = Nothing
    Set cmdSetHead = Nothing
    Set cmdSetShaft = Nothing
    Set lblBow = Nothing
    Set lblArrows = Nothing
    Set lblArrowHead = Nothing
    Set lblArrowShaft = Nothing
    Set chkEnableFletching = Nothing
    Set chkEnableBlacklist = Nothing
    Set chkEnableArcherBlacklist = Nothing
    
    'Melee tab
    Set txtMeleeRadius = Nothing
    Set cmdSetWeapon = Nothing
    Set cmdSetShield = Nothing
    Set lblWeapon = Nothing
    Set lblShield = Nothing
    
    'Nav tab
    Set chkEnableNav = Nothing
    Set cmdShowNavEditor = Nothing
    Set chkEnableSticky = Nothing
    Set cmdSetSticky = Nothing
    Set chkEnableFollow = Nothing
    Set cmdSetFollow = Nothing
    
    'Options
    Set chMacroSpellsLevel = Nothing
    Set chHealSpellsLevel = Nothing
    Set chEmergencySpellsLevel = Nothing
    
    Set chkLogoutOnTapers = Nothing
    Set chkLogoutOnKits = Nothing
    Set chkLogoutOnScarabs = Nothing
    Set chkLogoutOnArrows = Nothing
    Set chkEnableTurbo = Nothing
    Set chkEnableDOT = Nothing
    'Set sldMacroTimer = Nothing
    
    'Fellow
    Set chkEnableAutoFellow = Nothing
    Set txtFellowPassword = Nothing
    Set chkAutoFellowList = Nothing
    Set FellowList = Nothing
    
    'Misc
    Set chAttackHeight = Nothing
    Set chkUseManaCharge = Nothing
    Set chkUseManaStone = Nothing
    Set chkUseMassiveCharges = Nothing
    Set chkRechargeManaStones = Nothing
    'Set headerButton = Nothing
    Set btnStopSounds = Nothing
    Set btnResetHUD = Nothing
    Set btnExportInv = Nothing
            
    Unload = True
    MyDebug InterfaceName & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage "(interface:" & InterfaceName & ") Unload - " & Err.Description
    Resume Fin
End Function

Public Function UpdateDefaultDmg(ByVal iDmg As Integer, Optional bSilent As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If iDmg = DMG_NONE Then 'Autoselect
        iDmg = g_Spells.FindAndSetDefaultDamage
        If iDmg = DMG_NONE Then
            PrintErrorMessage "Could not find ANY valid default lvl " & g_Data.mMaxWarLevel & " War and/or lvl " & g_Data.mMaxVulnLevel & " Vuln spells for " & GetDamageString(iDmg) & " default damage. You can fix this by raising your spell levels, or learning the appropriate spells."
            bRet = False
        End If
    Else
        bRet = g_Spells.SetDefaultDamage(iDmg, "UpdateDefaultDmg")
        If Not bRet Then
            PrintErrorMessage "Could not find valid Default lvl " & g_Data.mMaxWarLevel & " War and/or lvl " & g_Data.mMaxVulnLevel & " Vuln spells for " & GetDamageString(iDmg) & " default damage. Please select another one or the macro may not work correctly."
        End If
    End If
        
    If bRet And Not bSilent Then
        PrintMessage "Default damage set to " & GetDamageString(iDmg)
    End If
    
Fin:
    UpdateDefaultDmg = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsUIMacro.UpdateDefaultDmg - " & Err.Description
    Resume Fin
End Function

'
' WithEvents Change Functions
'

'Private Sub txtMeleeRadius_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtMeleeRadius.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtMeleeRadius_Change - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub txtArcherRadius_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtArcherRadius.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtArcherRadius_Change - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub txtMageRadius_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtMageRadius.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtMageRadius_Change - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub txtVulnRange_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtVulnRange.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtVulnRange_Change - " & Err.Description
'    Resume Fin
'End Sub


'Private Sub txtArcRange_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtArcRange.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtArcRange_Change - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub txtRingRange_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtRingRange.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtRingRange_Change - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub txtRingNum_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtRingNum.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtRingNum_Change - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub txtAssistRange_Change(ByVal nID As Long, ByVal nText As String)
'On Error GoTo ErrorHandler
'
'    txtAssistRange.Text = Val(nText)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "clsUIMacro.txtAssistRange_Change - " & Err.Description
'    Resume Fin
'End Sub


Private Sub chDefaultDmg_Change(ByVal nID As Long, ByVal nIndex As Long)
On Error GoTo ErrorHandler

    If nIndex >= 0 Then
        Dim iDmg As eDamageType
        Dim sDmgName As String
        sDmgName = chDefaultDmg.Text(nIndex)
        iDmg = GetDamageType(sDmgName)
        Call UpdateDefaultDmg(iDmg)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMacro.chDefaultDmg_Change - " & Err.Description
    Resume Fin
End Sub

Private Sub chVulnLevel_Change(ByVal nID As Long, ByVal nIndex As Long)
    If nIndex > 0 Then
        g_Data.mMaxVulnLevel = nIndex + 1
        
        'Check if we get a valid default vuln spell
        Call UpdateDefaultDmg(g_Spells.DefaultDamage, True)
    End If
End Sub

Private Sub chWarLevel_Change(ByVal nID As Long, ByVal nIndex As Long)
    If nIndex > 0 Then
        g_Data.mMaxWarLevel = nIndex + 1
        
        'Check if we get a valid default war spell
        Call UpdateDefaultDmg(g_Spells.DefaultDamage, True)
    End If
End Sub

Private Sub Class_Terminate()
    Call Unload
End Sub

Private Function GetInventorySelection(ByVal sItemDescription As String, Optional ByRef objItemOut As acObject) As Long
On Error GoTo ErrorHandler

    Dim lRet As Long
    Dim lSel As Long

    lRet = 0
    lSel = g_Hooks.CurrentSelection
    If lSel <> 0 Then
        If Not g_Objects.Items.Inv.Exists(lSel) Then
            PrintMessage "Please select a valid " & sItemDescription & " in your inventory."
            Set objItemOut = Nothing
        Else
            Set objItemOut = g_Objects.Items.Inv(lSel)
            lRet = lSel
        End If
    Else
        PrintMessage "Please select a valid " & sItemDescription & " in your inventory."
        Set objItemOut = Nothing
    End If
    
Fin:
    GetInventorySelection = lRet
    Exit Function
ErrorHandler:
    lRet = 0
    Set objItemOut = Nothing
    PrintErrorMessage InterfaceName & ".GetInventorySelection - " & Err.Description
    Resume Fin
End Function

Private Sub UpdateItemLabel(ByRef objItem As acObject, ByRef lblControl As DecalControls.StaticText)
    If Valid(objItem) Then
        lblControl.Text = objItem.Name
    Else
        lblControl.Text = "Not Set"
    End If
End Sub

Public Sub UpdateMacroItemsLabels()
On Error GoTo ErrorHandler

    Call UpdateItemLabel(g_Data.Wand, lblWand)
    Call UpdateItemLabel(g_Data.Weapon, lblWeapon)
    Call UpdateItemLabel(g_Data.Shield, lblShield)
    Call UpdateItemLabel(g_Data.Bow, lblBow)
    
    Call UpdateItemLabel(g_Data.WeapBludge, lblWeapBludge)
    Call UpdateItemLabel(g_Data.WeapSlash, lblWeapSlash)
    Call UpdateItemLabel(g_Data.WeapPierce, lblWeapPierce)
    Call UpdateItemLabel(g_Data.WeapFire, lblWeapFire)
    Call UpdateItemLabel(g_Data.WeapCold, lblWeapCold)
    Call UpdateItemLabel(g_Data.WeapAcid, lblWeapAcid)
    Call UpdateItemLabel(g_Data.WeapLight, lblWeapLight)
    
    lblArrows.Text = g_Data.Arrows
    lblArrowHead.Text = g_Data.ArrowHead
    lblArrowShaft.Text = g_Data.ArrowShaft
    
    lblArrowBludge.Text = g_Data.m_sArrowBludge
    lblHeadBludge.Text = g_Data.m_sHeadBludge
    lblArrowSlash.Text = g_Data.m_sArrowSlash
    lblHeadSlash.Text = g_Data.m_sHeadSlash
    lblArrowPierce.Text = g_Data.m_sArrowPierce
    lblHeadPierce.Text = g_Data.m_sHeadPierce
    lblArrowFire.Text = g_Data.m_sArrowFire
    lblHeadFire.Text = g_Data.m_sHeadFire
    lblArrowCold.Text = g_Data.m_sArrowCold
    lblHeadCold.Text = g_Data.m_sHeadCold
    lblArrowAcid.Text = g_Data.m_sArrowAcid
    lblHeadAcid.Text = g_Data.m_sHeadAcid
    lblArrowLight.Text = g_Data.m_sArrowLight
    lblHeadLight.Text = g_Data.m_sHeadLight
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage InterfaceName & ".UpdateMacroItemsLabels - " & Err.Description
    Resume Fin
End Sub

Private Sub btnStopSounds_Accepted(ByVal nID As Long)
    'Stop all sounds
    Call StopLoopingSound
End Sub

Private Sub btnResetHUD_Accepted(ByVal bID As Long)
    'Reset the HUD positions
    Call g_HUD.SetMacroHudPos(20, 60)
    Call g_HUD.SetStatsHudPos(20, 120)
    Call g_HUD.SetInvHudPos(200, 60)
    Call g_HUD.SetDebugHudPos(20, 300)
    Call g_HUD.SetDOTHudPos(200, 200)
End Sub

Private Sub btnExportInv_Accepted(ByVal bID As Long)
    'Export Inventory
    Call g_Macro.Loot.exportInventory
End Sub

'Arrows for the archer tab
Private Sub cmdArrowBludge_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowBludge = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdArrowSlash_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowSlash = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdArrowPierce_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowPierce = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdArrowFire_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowFire = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdArrowCold_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowCold = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdArrowAcid_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowAcid = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdArrowLight_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.m_sArrowLight = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub

Private Sub cmdHeadBludge_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadBludge = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdHeadSlash_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadSlash = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdHeadPierce_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadPierce = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdHeadFire_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadFire = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdHeadCold_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadCold = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdHeadAcid_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadAcid = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdHeadLight_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Head", objItem)
    If Valid(objItem) Then
        g_Data.m_sHeadLight = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub

'Default arrows
Private Sub cmdSetArrow_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow", objItem)
    If Valid(objItem) Then
        g_Data.Arrows = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub

'Default heads for fletching
Private Sub cmdSetHead_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow Head", objItem)
    If Valid(objItem) Then
        g_Data.ArrowHead = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub

'Shafts for fletching
Private Sub cmdSetShaft_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Arrow Shaft", objItem)
    If Valid(objItem) Then
        g_Data.ArrowShaft = objItem.Name
        Call UpdateMacroItemsLabels
    End If
End Sub

'Default bow
Private Sub cmdSetBow_Accepted(ByVal nID As Long)
    g_Data.BowGUID = GetInventorySelection("Bow")
    Call UpdateItemLabel(g_Data.Bow, lblBow)
End Sub

'Default melee Weapon
Private Sub cmdSetWeapon_Accepted(ByVal nID As Long)
    g_Data.WeaponGUID = GetInventorySelection("Melee Weapon")
    Call UpdateItemLabel(g_Data.Weapon, lblWeapon)
End Sub

'Melee Shield
Private Sub cmdSetShield_Accepted(ByVal nID As Long)
    g_Data.ShieldGUID = GetInventorySelection("Shield")
    Call UpdateItemLabel(g_Data.Shield, lblShield)
End Sub

'Default wand
Private Sub cmdSetWand_Accepted(ByVal nID As Long)
    g_Data.WandGUID = GetInventorySelection("Wand")
    Call UpdateItemLabel(g_Data.Wand, lblWand)
End Sub

'Elemental Weapons (could be wand, melee or archer)
Private Sub cmdWeapBludge_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        'MyDebug "clsUIMacro.cmdWeapBludge: " & objItem.Name
        g_Data.SetWeapBludgeGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdWeapSlash_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        g_Data.SetWeapSlashGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdWeapPierce_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        g_Data.SetWeapPierceGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdWeapFire_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        g_Data.SetWeapFireGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdWeapCold_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        g_Data.SetWeapColdGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdWeapAcid_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        g_Data.SetWeapAcidGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub
Private Sub cmdWeapLight_Accepted(ByVal nID As Long)
    Dim objItem As acObject
    Call GetInventorySelection("Weapon", objItem)
    If Valid(objItem) Then
        g_Data.SetWeapLightGUID = objItem.Guid
        Call UpdateMacroItemsLabels
    End If
End Sub


' Navigation
Private Sub cmdShowNavEditor_Accepted(ByVal nID As Long)
    Call g_ui.NavEditor.ShowEditor
End Sub

Private Sub chkEnableNav_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If (bChecked) Then
        ' Make sure Sticky is not enabled
        chkEnableSticky.Checked = False
        chkEnableFollow.Checked = False
    End If
End Sub

Private Sub chkEnableSticky_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If (bChecked) Then
        MyDebug "chkEnableSticky_Change: on : " & bChecked
        ' Make sure Nav is not enabled
        If (chkEnableNav.Checked) Then
            chkEnableNav.Checked = False
        End If
        If (chkEnableFollow.Checked) Then
            chkEnableFollow.Checked = False
        End If
    Else
        MyDebug "chkEnableSticky_Change: off : " & bChecked
        Call g_Nav.SetStickyOff
    End If
End Sub

Private Sub cmdSetSticky_accepted(ByVal nID As Long)
    ' Set the sticky position
    If chkEnableSticky.Checked Then
        Call g_Nav.SetSticky
    Else
        If (chkEnableNav.Checked) Then
            chkEnableNav.Checked = False
        End If
        If (chkEnableFollow.Checked) Then
            chkEnableFollow.Checked = False
        End If
        chkEnableSticky.Checked = True
        Call g_Nav.SetSticky
    End If

End Sub

Private Sub chkEnableFollow_Change(ByVal nID As Long, ByVal bChecked As Boolean)
    If (bChecked) Then
        MyDebug "chkEnableFollower_Change: on : " & bChecked
        ' Make sure Nav is not enabled
        If (chkEnableNav.Checked) Then
            chkEnableNav.Checked = False
        End If
        If (chkEnableSticky.Checked) Then
            chkEnableSticky.Checked = False
            Call g_Nav.SetStickyOff
        End If
    Else
        MyDebug "chkEnableFollower_Change: off : " & bChecked
        Call g_Nav.SetFollowOff
    End If
End Sub

Private Sub cmdSetFollow_accepted(ByVal nID As Long)
    ' Set the sticky position
    If chkEnableFollow.Checked Then
        Call g_Nav.SetFollow
    Else
        If (chkEnableNav.Checked) Then
            chkEnableNav.Checked = False
        End If
        If (chkEnableSticky.Checked) Then
            chkEnableSticky.Checked = False
            Call g_Nav.SetStickyOff
        End If
        chkEnableFollow.Checked = True
        Call g_Nav.SetFollow
    End If

End Sub


' Fellow list changes
Private Sub FellowList_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Select Case nX
    Case 0
        Call g_FellowList.removeFromLine(FellowList.Data(1, nY))
    Case 1
        If (g_Hooks.ChatState = False) Then SendKeys "{ENTER}"
        SendKeys "@tell " & FellowList.Data(1, nY) & ", "
    Case 2
        Call g_FellowList.moveUp(nY + 1)
    Case 3
        Call g_FellowList.moveDown(nY + 1)
    End Select
 
End Sub

'Private Sub sldMacroTimer_Change(ByVal nID As Long, ByVal lValue As Long)
'On Error GoTo Error_Handler
'
'    g_Core.TimerInterval = lValue
'
'Fin:
'    Exit Sub
'Error_Handler:
'    PrintErrorMessage "clsUIMain.sldMacroTimer_Change - " & Err.Description
'    Resume Fin
'End Sub




' -----------------------------
' For new Healing/Restam list stuff
' Sub to add the selected item to the list of items used for healing.
' This sub first checks to see if the item is already on the list, if
' it is, then the sub does nothing, otherwise it adds it to the end
' of the list.  Items that are added to the list are limited to
' ITEM_HEALING_KIT or ITEM_FOOD types

Private Sub btnAddStamItem_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    Dim aObj As acObject
    Dim aName As String
        
    If (g_Hooks.CurrentSelection <> 0) Then
        If g_Objects.Items.Exists(g_Hooks.CurrentSelection) Then
            Set aObj = g_Objects.Items(g_Hooks.CurrentSelection)

                Dim intFoundItem As Integer
                intFoundItem = 0
                Dim intTempCounter As Integer
                For intTempCounter = 0 To lstStamItems.Count - 1
                    If (lstStamItems.Data(1, intTempCounter) = aObj.Name) Then
                        intFoundItem = 1
                        Exit For
                    End If
                Next
                If (intFoundItem = 0) Then
                    intTempCounter = lstStamItems.AddRow
                    lstStamItems.Data(COL_ENABLED, intTempCounter) = True
                    lstStamItems.Data(COL_ITEM_NAME, intTempCounter) = aObj.Name
                    lstStamItems.Data(COL_ICON_UP, intTempCounter, 1) = MOVE_UP_ICON
                    lstStamItems.Data(COL_ICON_DOWN, intTempCounter, 1) = MOVE_DOWN_ICON
                    Call SaveItemsList(lstStamItems, GetStamListPath)
                    PrintMessage "Adding: " & aObj.Name
                End If
        End If
    End If

Fin:
    Set aObj = Nothing
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIMain.btnAddStamItem_Accepted - " & Err.Description
    Resume Fin
End Sub


Private Sub lstStamItems_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Select Case nX
    Case 0
        'Enable checkbox
    Case 1
        'Delete this item
        PrintMessage "Removing: " & lstStamItems.Data(COL_ITEM_NAME, nY)
        lstStamItems.DeleteRow (nY)
        Call SaveItemsList(lstStamItems, GetStamListPath)
    Case 2
        ' Move up the list
        Call lstMoveUp(lstStamItems, nY)
    Case 3
        ' Move down the list
        Call lstMoveDown(lstStamItems, nY)
    End Select
 
End Sub


' Healing items
Private Sub btnAddHealItem_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    Dim aObj As acObject
    Dim aName As String
        
    If (g_Hooks.CurrentSelection <> 0) Then
        If g_Objects.Items.Exists(g_Hooks.CurrentSelection) Then
            Set aObj = g_Objects.Items(g_Hooks.CurrentSelection)

            If (aObj.itemType = ITEM_HEALING_KIT) Or (aObj.itemType = ITEM_FOOD) Then
                Dim intFoundItem As Integer
                intFoundItem = 0
                Dim intTempCounter As Integer
                For intTempCounter = 0 To lstHealItems.Count - 1
                    If (lstHealItems.Data(1, intTempCounter) = aObj.Name) Then
                        intFoundItem = 1
                        Exit For
                    End If
                Next
                If (intFoundItem = 0) Then
                    intTempCounter = lstHealItems.AddRow
                    lstHealItems.Data(COL_ENABLED, intTempCounter) = True
                    lstHealItems.Data(COL_ITEM_NAME, intTempCounter) = aObj.Name
                    lstHealItems.Data(COL_ICON_UP, intTempCounter, 1) = MOVE_UP_ICON
                    lstHealItems.Data(COL_ICON_DOWN, intTempCounter, 1) = MOVE_DOWN_ICON
                    Call SaveItemsList(lstHealItems, GetHealListPath)
                    PrintMessage "Adding: " & aObj.Name
                End If
            End If
        End If
    End If

Fin:
    Set aObj = Nothing
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIMain.btnAddHealItem_Accepted - " & Err.Description
    Resume Fin
End Sub


Private Sub lstHealItems_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Select Case nX
    Case 0
        'Enable checkbox
    Case 1
        'Delete this item
        PrintMessage "Removing: " & lstHealItems.Data(COL_ITEM_NAME, nY)
        lstHealItems.DeleteRow (nY)
        Call SaveItemsList(lstHealItems, GetHealListPath)
    Case 2
        ' Move up the list
        Call lstMoveUp(lstHealItems, nY)
    Case 3
        ' Move down the list
        Call lstMoveDown(lstHealItems, nY)
    End Select
 
End Sub


' Emergency items
Private Sub btnAddEmergItem_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler

    Dim aObj As acObject
    Dim aName As String
        
    If (g_Hooks.CurrentSelection <> 0) Then
        If g_Objects.Items.Exists(g_Hooks.CurrentSelection) Then
            Set aObj = g_Objects.Items(g_Hooks.CurrentSelection)

            Dim intTempCounter As Integer
            Dim intFoundItem As Integer
            
            intFoundItem = 0
            For intTempCounter = 0 To lstEmergItems.Count - 1
                If (lstEmergItems.Data(1, intTempCounter) = aObj.Name) Then
                    intFoundItem = 1
                    Exit For
                End If
            Next
            If (intFoundItem = 0) Then
                intTempCounter = lstEmergItems.AddRow
                lstEmergItems.Data(COL_ENABLED, intTempCounter) = True
                lstEmergItems.Data(COL_ITEM_NAME, intTempCounter) = aObj.Name
                lstEmergItems.Data(COL_ICON_UP, intTempCounter, 1) = MOVE_UP_ICON
                lstEmergItems.Data(COL_ICON_DOWN, intTempCounter, 1) = MOVE_DOWN_ICON
                Call SaveItemsList(lstEmergItems, GetEmergListPath)
                PrintMessage "Adding: " & aObj.Name
            End If
        End If
    End If

Fin:
    Set aObj = Nothing
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsUIMain.btnAddEmergItem_Accepted - " & Err.Description
    Resume Fin
End Sub


Private Sub lstEmergItems_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
    Select Case nX
    Case 0
        'Enable checkbox
    Case 1
        'Delete this item
        PrintMessage "Removing: " & lstEmergItems.Data(COL_ITEM_NAME, nY)
        lstEmergItems.DeleteRow (nY)
        Call SaveItemsList(lstEmergItems, GetEmergListPath)
    Case 2
        ' Move up the list
        Call lstMoveUp(lstEmergItems, nY)
    Case 3
        ' Move down the list
        Call lstMoveDown(lstEmergItems, nY)
    End Select
 
End Sub

'Swap with the item before it in the list (move UP in line)
Public Sub lstMoveUp(ByVal lstControl As DecalControls.list, ByVal idx As Long)
On Error GoTo ErrorHandler
    
    If (idx = 0) Then Exit Sub
    
    Dim tmpName As String
    Dim tmpBool As Boolean

    tmpName = lstControl.Data(COL_ITEM_NAME, idx)
    tmpBool = lstControl.Data(COL_ENABLED, idx)
    
    lstControl.Data(COL_ITEM_NAME, idx) = lstControl.Data(COL_ITEM_NAME, idx - 1)
    lstControl.Data(COL_ITEM_NAME, idx - 1) = tmpName
    
    lstControl.Data(COL_ENABLED, idx) = lstControl.Data(COL_ENABLED, idx - 1)
    lstControl.Data(COL_ENABLED, idx - 1) = tmpBool
    
Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMacro.lstMoveUp: " & Err.Description
End Sub

'Swap with the item after it in the list (move DOWN in line)
Public Sub lstMoveDown(ByVal lstControl As DecalControls.list, ByVal idx As Long)
On Error GoTo ErrorHandler
    
    If (idx >= lstControl.Count) Then Exit Sub
    
    Dim tmpName As String
    Dim tmpBool As Boolean

    tmpName = lstControl.Data(COL_ITEM_NAME, idx)
    tmpBool = lstControl.Data(COL_ENABLED, idx)
    
    lstControl.Data(COL_ITEM_NAME, idx) = lstControl.Data(COL_ITEM_NAME, idx + 1)
    lstControl.Data(COL_ITEM_NAME, idx + 1) = tmpName
    
    lstControl.Data(COL_ENABLED, idx) = lstControl.Data(COL_ENABLED, idx + 1)
    lstControl.Data(COL_ENABLED, idx + 1) = tmpBool
    
Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMacro.lstMoveDown: " & Err.Description
End Sub

Private Function GetHealListPath() As String
    GetHealListPath = g_Settings.Profile.FullPath & "\" & FILE_HEALITEMS_LIST
End Function

Private Function GetStamListPath() As String
    GetStamListPath = g_Settings.Profile.FullPath & "\" & FILE_STAMITEMS_LIST
End Function

Private Function GetEmergListPath() As String
    GetEmergListPath = g_Settings.Profile.FullPath & "\" & FILE_EMERGITEMS_LIST
End Function

'Save the items list
Private Function SaveItemsList(ByRef lstControl As DecalControls.list, ByVal sFullPath As String) As Boolean
On Error GoTo Error_Handler
    
    Dim bRet As Boolean
    Dim db As New DataFile
    Dim dat As clsDataEntry
    
    Dim i As Integer
    
    MyDebug "clsUIMacro.SaveItemsList - Saving list to " & sFullPath
    
    For i = 0 To lstControl.Count - 1
        Set dat = New clsDataEntry
        dat.AddParam TAG_ENABLED, BoolToInteger(lstControl.Data(COL_ENABLED, i))
        dat.AddParam TAG_ITEM_NAME, lstControl.Data(COL_ITEM_NAME, i)
        
        If Not db.AddData(dat) Then
            PrintErrorMessage "clsUIMacro.SaveItemsList - db.AddData(dat) failed"
        End If
    Next i
    
    bRet = db.save(sFullPath)
    
    If Not bRet Then
        PrintErrorMessage "clsUIMacro.SaveItemsList - Failed to save list to " & sFullPath
    End If
    
Fin:
    Set db = Nothing
    Set dat = Nothing
    SaveItemsList = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage "clsUIMacro.SaveItemsList(" & sFullPath & ") - " & Err.Description
    Resume Fin
End Function

Private Function LoadItemsList(ByRef lstControl As DecalControls.list, ByVal sFullPath As String) As Boolean
On Error GoTo Error_Handler
    
    Dim bRet As Boolean
    Dim db As New DataFile
    Dim dat As clsDataEntry
    
    Dim i As Integer
    
    MyDebug "clsUIMacro.LoadItemsList - Loading list from " & sFullPath
    
    'first reset the list
    Call lstControl.Clear
    
    'Load the data
    If db.Load(sFullPath) Then
    
        Dim bEnable As Boolean
        Dim sItemName As String
        
        For Each dat In db
            If dat.ParamExist(TAG_ITEM_NAME) Then
                sItemName = dat.Param(TAG_ITEM_NAME)
                
                If dat.ParamExist(TAG_ENABLED) Then
                    bEnable = dat.Param(TAG_ENABLED)
                Else
                    bEnable = False
                End If
                
                i = lstControl.AddRow
                lstControl.Data(COL_ENABLED, i) = bEnable
                lstControl.Data(COL_ITEM_NAME, i) = sItemName
                lstControl.Data(COL_ICON_UP, i, 1) = MOVE_UP_ICON
                lstControl.Data(COL_ICON_DOWN, i, 1) = MOVE_DOWN_ICON
                    
            End If
        Next dat
    Else
        PrintErrorMessage "clsUIMacro.LoadItemsList - Failed to load " & sFullPath
    End If
    
Fin:
    Set db = Nothing
    Set dat = Nothing
    LoadItemsList = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage "clsUIMacro.LoadItemsList(" & sFullPath & ") - " & Err.Description
    Resume Fin
End Function

'Returns the first item found from the list
Public Function findObjectFromList(ByRef lstControl As DecalControls.list) As acObject
On Error GoTo ErrorMessage
          
    Dim aObj As acObject
    Dim curItemName As String
    Dim i As Integer
    
    ' Look thru the List and return first one found
    For i = 0 To lstControl.Count - 1
    
        If lstControl.Data(COL_ENABLED, i) Then
            curItemName = lstControl.Data(COL_ITEM_NAME, i)
            Set aObj = g_Objects.Items.InvFindByName(curItemName, False)
            If Valid(aObj) Then
                'MyDebug "clsUIMacro.findObjectFromList: Found item: " & aObj.Name
                GoTo Fin
            End If
        End If
        
    Next i
  
    If Not Valid(aObj) Then
        'MyDebug "clsUIMacro.findObjectFromList: Using Default: " & STR_ITEM_HEALTH_POTION
        Set aObj = g_Objects.Items.InvFindByName(STR_ITEM_HEALTH_POTION, False)
    End If
        
Fin:
    Set findObjectFromList = aObj
    Set aObj = Nothing
    Exit Function
ErrorMessage:
    Set aObj = Nothing
    PrintErrorMessage "clUIMacro.findObjectFromList - " & Err.Description
    Resume Fin
End Function

