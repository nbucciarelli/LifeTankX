VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mLogOutputPath As String
Public mDebugMode As Boolean

Public mdbMonsters As dbMonsters    'monsters database
Private m_LootFilters As clsLootFiltersManager
Private m_Exceptions As clsExceptions
Private m_AutoResponse As clsAutoResponse

Public memCurrentItem1 As Long
Public memCurrentItem2 As Long
Public memCastSpell As Long

Public mMaxWarLevel As Integer
Public mMaxVulnLevel As Integer
Public mMaxBuffLevel As Integer

'Macro equipment
Private m_WeaponGUID As Long
Private m_objWeapon As acObject
Private m_ShieldGUID As Long
Private m_objShield As acObject
Private m_BowGUID As Long
Private m_objBow As acObject
Private m_WandGUID As Long
Private m_objWand As acObject
    
Private m_sArrows As String
Private m_sArrowHead As String
Private m_sArrowShaft As String

Public m_sArrowBludge As String
Public m_sHeadBludge As String
Public m_sArrowSlash As String
Public m_sHeadSlash As String
Public m_sArrowPierce As String
Public m_sHeadPierce As String
Public m_sArrowFire As String
Public m_sHeadFire As String
Public m_sArrowCold As String
Public m_sHeadCold As String
Public m_sArrowAcid As String
Public m_sHeadAcid As String
Public m_sArrowLight As String
Public m_sHeadLight As String

Public WeapBludge As acObject
Public WeapBludgeGUID As Long
Public WeapSlash As acObject
Public WeapSlashGUID As Long
Public WeapPierce As acObject
Public WeapPierceGUID As Long
Public WeapFire As acObject
Public WeapFireGUID As Long
Public WeapCold As acObject
Public WeapColdGUID As Long
Public WeapAcid As acObject
Public WeapAcidGUID As Long
Public WeapLight As acObject
Public WeapLightGUID As Long


'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    mLogOutputPath = PATH_LOGS
    mDebugMode = True
    
    mMaxWarLevel = 8
    mMaxVulnLevel = 7
    mMaxBuffLevel = 8
    
    m_WeaponGUID = 0
    m_ShieldGUID = 0
    m_BowGUID = 0
    m_WandGUID = 0
    
    Set m_objBow = Nothing
    Set m_objShield = Nothing
    Set m_objWeapon = Nothing
    Set m_objWand = Nothing

    Set WeapBludge = Nothing
    WeapBludgeGUID = 0
    Set WeapSlash = Nothing
    WeapSlashGUID = 0
    Set WeapPierce = Nothing
    WeapPierceGUID = 0
    Set WeapFire = Nothing
    WeapFireGUID = 0
    Set WeapCold = Nothing
    WeapColdGUID = 0
    Set WeapAcid = Nothing
    WeapAcidGUID = 0
    Set WeapLight = Nothing
    WeapLightGUID = 0
    
    Set mdbMonsters = New dbMonsters
    Set m_LootFilters = New clsLootFiltersManager
    Set m_Exceptions = New clsExceptions
    Set m_AutoResponse = New clsAutoResponse
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDatas.Class_Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set mdbMonsters = Nothing
    Set m_objWeapon = Nothing
    Set m_objShield = Nothing
    Set m_objBow = Nothing
    Set m_Exceptions = Nothing
    Set m_LootFilters = Nothing
    Set m_AutoResponse = Nothing
End Sub

'#####################################################################################
'#
'#                                PROPERTIES
'#
'#####################################################################################
Public Property Get LootFilters() As clsLootFiltersManager
    Set LootFilters = m_LootFilters
End Property

Public Property Get Exceptions() As clsExceptions
    Set Exceptions = m_Exceptions
End Property

Public Property Get getClassAutoResponse() As clsAutoResponse
    Set getClassAutoResponse = m_AutoResponse
End Property

Public Property Get MinManaThreshold() As Integer
    MinManaThreshold = g_ui.Macro.sldMana.SliderPosition
End Property

Public Property Get MinStamThreshold() As Integer
    MinStamThreshold = g_ui.Macro.sldStam.SliderPosition
End Property

Public Property Get MinHealthThreshold() As Integer
    MinHealthThreshold = g_ui.Macro.sldHealth.SliderPosition
End Property

Public Property Get CriticalHealthThreshold() As Integer
    CriticalHealthThreshold = g_ui.Macro.sldCritHealth.SliderPosition
End Property

Public Property Get FellowHealthThreshold() As Integer
    FellowHealthThreshold = g_ui.Macro.sldFellowHealth.SliderPosition
End Property

Public Property Get FellowStamThreshold() As Integer
    FellowStamThreshold = g_ui.Macro.sldFellowStam.SliderPosition
End Property

Public Property Get SalvagerFrequency() As Integer
    SalvagerFrequency = Val(g_ui.Loot.txtSalvageFrequency.Text)
End Property

'Ranges
Public Property Get DangerZoneRadius() As Integer
    DangerZoneRadius = Val(g_ui.Macro.txtDangerZone.Text)
End Property

Public Property Get MacroSpellsLevel() As Integer
    MacroSpellsLevel = g_ui.Macro.chMacroSpellsLevel.Selected + 1
End Property

Public Property Get HealSpellsLevel() As Integer
    HealSpellsLevel = g_ui.Macro.chHealSpellsLevel.Selected + 1
End Property

Public Property Get EmergencySpellsLevel() As Integer
    EmergencySpellsLevel = g_ui.Macro.chEmergencySpellsLevel.Selected + 1
End Property

Public Property Get BuffsSpellsLevel() As Integer
    BuffsSpellsLevel = g_ui.Buffs.chMaxBuffSpellLevel.Selected + 1
End Property

Public Property Get RebuffInterval() As Integer
    RebuffInterval = Val(g_ui.Buffs.txtRebuffInterval.Text)
End Property

Public Property Get NumContinuousBuffs() As Integer
    NumContinuousBuffs = Val(g_ui.Buffs.txtNumContinuousBuffs.Text)
End Property

Public Property Get LootSearchRadius() As Integer
    LootSearchRadius = Val(g_ui.Loot.txtPickupRange.Text)
End Property

Public Property Get LootMinValue() As Long
    LootMinValue = Val(g_ui.Loot.txtPickupMinValue.Text)
End Property

Public Property Get HighManaValue() As Integer
    HighManaValue = Val(g_ui.Loot.txtHighMana.Text)
End Property

Public Property Get LootMinBurdenRatio() As Integer
    LootMinBurdenRatio = Val(g_ui.Loot.txtBurdenRatio.Text)
End Property

Public Property Get NumToPickup() As Integer
    NumToPickup = Val(g_ui.Loot.txtNumToPickup.Text)
End Property

Public Property Get CastBreakerTime() As Integer    'in Seconds!
    'CastBreakerTime = g_ui.Macro.sldCastBreaker.SliderPosition
    CastBreakerTime = 1
End Property

'------------------------------------
'Macro Equipments
'------------------------------------
'GET Objects
Public Property Get Weapon() As acObject
    Set Weapon = m_objWeapon
End Property

Public Property Get Shield() As acObject
    Set Shield = m_objShield
End Property

Public Property Get Bow() As acObject
    Set Bow = m_objBow
End Property

Public Property Get Wand() As acObject
    Set Wand = m_objWand
End Property

'Fletching
Public Property Get Arrows() As String
    Arrows = m_sArrows
End Property

Public Property Let Arrows(ByVal sVal As String)
    m_sArrows = sVal
End Property

Public Property Get ArrowHead() As String
    ArrowHead = m_sArrowHead
End Property

Public Property Let ArrowHead(ByVal sVal As String)
    m_sArrowHead = sVal
End Property

Public Property Get ArrowShaft() As String
    ArrowShaft = m_sArrowShaft
End Property

Public Property Let ArrowShaft(ByVal sVal As String)
    m_sArrowShaft = sVal
End Property

'GET GUIDs
Public Property Get WeaponGUID() As Long
    WeaponGUID = m_WeaponGUID
End Property

Public Property Get ShieldGUID() As Long
    ShieldGUID = m_ShieldGUID
End Property

Public Property Get BowGUID() As Long
    BowGUID = m_BowGUID
End Property

Public Property Get WandGUID() As Long
    WandGUID = m_WandGUID
End Property

'SET GUIDs
Public Property Let WeaponGUID(ByVal lVal As Long)
    Call SetGUID(lVal, m_WeaponGUID, m_objWeapon)
End Property

Public Property Let ShieldGUID(ByVal lVal As Long)
    Call SetGUID(lVal, m_ShieldGUID, m_objShield)
End Property

Public Property Let BowGUID(ByVal lVal As Long)
    Call SetGUID(lVal, m_BowGUID, m_objBow)
End Property

Public Property Let WandGUID(ByVal lVal As Long)
    Call SetGUID(lVal, m_WandGUID, m_objWand)
End Property

Public Property Let SetWeapBludgeGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapBludgeGUID, WeapBludge)
End Property
Public Property Let SetWeapSlashGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapSlashGUID, WeapSlash)
End Property
Public Property Let SetWeapPierceGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapPierceGUID, WeapPierce)
End Property
Public Property Let SetWeapFireGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapFireGUID, WeapFire)
End Property
Public Property Let SetWeapColdGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapColdGUID, WeapCold)
End Property
Public Property Let SetWeapAcidGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapAcidGUID, WeapAcid)
End Property
Public Property Let SetWeapLightGUID(ByVal lVal As Long)
    Call SetGUID(lVal, WeapLightGUID, WeapLight)
End Property


'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

Private Sub SetGUID(ByVal lVal As Long, ByRef lGUID As Long, ByRef objItem As acObject)
    lGUID = lVal
    
    If lVal = 0 Then Exit Sub
    
    If Valid(g_Objects) Then
        If g_Objects.Items.Inv.Exists(lVal) Then
            Set objItem = g_Objects.Items.Inv(lVal)
        Else
            MyDebug "clsDatas.SetGUID: lVal does not exist: " & lVal
            Set objItem = Nothing
        End If
    Else
        PrintErrorMessage "clsDatas.SetGUID : g_objects is NULL"
        Set objItem = Nothing
    End If
End Sub

'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################

'Public Sub GetMemlocks()
'    MyDebug "Retreiving memlocs from xml for castspell function"
'    memCurrentItem1 = g_PluginSite.QueryMemLoc("CurrentItem1")
'    memCurrentItem2 = g_PluginSite.QueryMemLoc("CurrentItem2")
'    memCastSpell = g_PluginSite.QueryMemLoc("CastSpell")
'    MyDebug ".....CurrentItem1 = " & memCurrentItem1 & " [Hex=" & Hex(memCurrentItem1) & "]"
'    MyDebug ".....CurrentItem2 = " & memCurrentItem2 & " [Hex=" & Hex(memCurrentItem2) & "]"
'    MyDebug ".....CastSpell = " & memCastSpell & " [Hex=" & Hex(memCastSpell) & "]"
'End Sub

Public Function GetWeaponType(ByVal iType As Integer) As acObject
On Error GoTo ErrorHandler
    
    Dim aWeap As acObject
    
    Set aWeap = Nothing
        
    Select Case iType
        Case DMG_SLASHING
            If g_ui.Macro.chkWeapSlash.Checked Then
                Set aWeap = g_Data.WeapSlash
            End If
        Case DMG_BLUDGEONING
            If g_ui.Macro.chkWeapBludge.Checked Then
                Set aWeap = g_Data.WeapBludge
            End If
        Case DMG_PIERCING
            If g_ui.Macro.chkWeapPierce.Checked Then
                Set aWeap = g_Data.WeapPierce
            End If
        Case DMG_FIRE
            If g_ui.Macro.chkWeapFire.Checked Then
                Set aWeap = g_Data.WeapFire
            End If
        Case DMG_COLD
            If g_ui.Macro.chkWeapCold.Checked Then
                Set aWeap = g_Data.WeapCold
            End If
        Case DMG_ACID
            If g_ui.Macro.chkWeapAcid.Checked Then
                Set aWeap = g_Data.WeapAcid
            End If
        Case DMG_LIGHTNING
            If g_ui.Macro.chkWeapLight.Checked Then
                Set aWeap = g_Data.WeapLight
            End If
    End Select

    Set GetWeaponType = aWeap

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsDatas.GetWeaponType - " & Err.Description
    Resume Fin
End Function

Public Function GetArrowType(Optional ByVal ooaCheck As Boolean = True) As String
On Error GoTo ErrorHandler

    Dim ArrowName As String
    Dim weapName As String
    Dim NumArrowHeads As Integer
    
    ArrowName = TAG_NOT_SET
    
    If Valid(g_currentEquip) Then
        weapName = g_currentEquip.Name
    End If
    
    MyDebug "clsDatas.GetArrowType: weapon: :" & weapName & ":"

    If g_ui.Macro.chkWeapSlash.Checked And Valid(g_Data.WeapSlash) Then
        If (weapName = g_Data.WeapSlash.Name) Then
            ArrowName = m_sArrowSlash
        End If
        End If
    If g_ui.Macro.chkWeapBludge.Checked And Valid(g_Data.WeapBludge) Then
        If (weapName = g_Data.WeapBludge.Name) Then
            ArrowName = m_sArrowBludge
        End If
        End If
    If g_ui.Macro.chkWeapPierce.Checked And Valid(g_Data.WeapPierce) Then
        If (weapName = g_Data.WeapPierce.Name) Then
            ArrowName = m_sArrowPierce
        End If
    End If
    If g_ui.Macro.chkWeapFire.Checked And Valid(g_Data.WeapFire) Then
        MyDebug "clsData.GetArrowType: FireWeap: :" & g_Data.WeapFire.Name & ":"
        If (weapName = g_Data.WeapFire.Name) Then
            ArrowName = m_sArrowFire
        End If
    End If
    If g_ui.Macro.chkWeapCold.Checked And Valid(g_Data.WeapCold) Then
        If (weapName = g_Data.WeapCold.Name) Then
            ArrowName = m_sArrowCold
        End If
    End If
    If g_ui.Macro.chkWeapAcid.Checked And Valid(g_Data.WeapAcid) Then
        If (weapName = g_Data.WeapAcid.Name) Then
            ArrowName = m_sArrowAcid
        End If
    End If
    If g_ui.Macro.chkWeapLight.Checked And Valid(g_Data.WeapLight) Then
        If (weapName = g_Data.WeapLight.Name) Then
            ArrowName = m_sArrowLight
        End If
    End If
    
    If ArrowName = TAG_NOT_SET Then
        MyDebug "GetArrowType: arrow not set for this bow. Using Default"
        ArrowName = m_sArrows
    End If
    
    GetArrowType = ArrowName
    
    MyDebug "clsDatas.GetArrowType: " & ArrowName
    
    NumArrowHeads = g_Objects.Items.InvCntByName(ArrowName, True)
    
    If NumArrowHeads > 0 Then
        MyDebug "clsDatas.GetArrowType: Found a total of " & NumArrowHeads & " " & ArrowName & " in inventory"
        Set g_currentArrow = g_Objects.Items.InvFindByName(ArrowName)
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsDatas.GetArrowType - " & Err.Description
    GetArrowType = m_sArrows
    Resume Fin
End Function

'Arrow heads for fletching
Public Function GetArrowHead(Optional ByVal ooaCheck As Boolean = True) As String
On Error GoTo ErrorHandler

    Dim ArrowHeadName As String
    Dim NumArrowHeads As Integer
    Dim weapName As String
    
    ArrowHeadName = TAG_NOT_SET
    
    If Valid(g_currentEquip) Then
        weapName = g_currentEquip.Name
    End If
    
    If g_ui.Macro.chkWeapSlash.Checked Then
        If (weapName = g_Data.WeapSlash.Name) Then
            ArrowHeadName = m_sHeadSlash
        End If
    End If
    If g_ui.Macro.chkWeapBludge.Checked Then
        If (weapName = g_Data.WeapBludge.Name) Then
            ArrowHeadName = m_sHeadBludge
        End If
    End If
    If g_ui.Macro.chkWeapPierce.Checked Then
        If (weapName = g_Data.WeapPierce.Name) Then
            ArrowHeadName = m_sHeadPierce
        End If
    End If
    If g_ui.Macro.chkWeapFire.Checked Then
        If (weapName = g_Data.WeapFire.Name) Then
            ArrowHeadName = m_sHeadFire
        End If
    End If
    If g_ui.Macro.chkWeapCold.Checked Then
        If (weapName = g_Data.WeapCold.Name) Then
            ArrowHeadName = m_sHeadCold
        End If
    End If
    If g_ui.Macro.chkWeapAcid.Checked Then
        If (weapName = g_Data.WeapAcid.Name) Then
            ArrowHeadName = m_sHeadAcid
        End If
    End If
    If g_ui.Macro.chkWeapLight.Checked Then
        If (weapName = g_Data.WeapLight.Name) Then
            ArrowHeadName = m_sHeadLight
        End If
    End If
    
    If ArrowHeadName = TAG_NOT_SET Then
        MyDebug "GetArrowHead: arrow Head not set for this type of arrow"
        ArrowHeadName = m_sArrowHead
    End If
    
    GetArrowHead = ArrowHeadName
    
    MyDebug "GetArrowHead: " & ArrowHeadName
    
    NumArrowHeads = g_Objects.Items.InvCntByName(ArrowHeadName, True)
    
    If NumArrowHeads > 0 Then
        MyDebug "GetArrowHead: Found a total of " & NumArrowHeads & " " & ArrowHeadName & " in inventory"
        'Set g_currentArrow = g_Objects.Items.InvFindByName(ArrowHeadName)
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsDatas.GetArrowHead - " & Err.Description
    Resume Fin
End Function


'Set debug mode
Public Sub SetDebugMode(ByVal bVal As Boolean)
    mDebugMode = bVal
    Call g_Filters.dsFilter.SetDebugMode(bVal)
End Sub

' Load and Save Functions
Public Function LoadMonsters(Optional ByVal sBaseFolder As String = "") As Boolean
    If sBaseFolder = "" Then
        sBaseFolder = g_Settings.GetDataFolder & "\" & PATH_DATA
    End If
    
    LoadMonsters = mdbMonsters.LoadDatabase(sBaseFolder & "\" & FILE_MONSTERS)
End Function

Public Function LoadLootFilters(Optional ByVal sBaseFolder As String = "") As Boolean
    If sBaseFolder = "" Then
        sBaseFolder = g_Settings.GetDataFolder & "\" & PATH_DATA
    End If
    
    LoadLootFilters = m_LootFilters.LoadFilters(sBaseFolder & "\" & FILE_LOOT_FILTERS)
End Function

Public Function SaveLootFilters(Optional ByVal sPath As String = "") As Boolean
    SaveLootFilters = m_LootFilters.SaveFilters(sPath)
End Function



