VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLootFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEBUG_ME = False

Private Const TAG_FILTER_ID = "id"
Private Const TAG_ARMOR_TYPE = "armortype"
Private Const TAG_ARMOR_COVERAGE = "coverage"
Private Const TAG_ARMOR_MIN_AL = "minal"

Private Const TAG_WEAPON_TYPE = "weaptype"
Private Const TAG_WEAPON_ELEMENT = "elem"
Private Const TAG_WEAPON_MIN_LOW_DMG = "minlowdmg"
Private Const TAG_WEAPON_MIN_HIGH_DMG = "minhighdmg"
Private Const TAG_WEAPON_MIN_DMG_MOD = "mindmgmod"
Private Const TAG_WEAPON_ATT_BONUS = "atk"
Private Const TAG_WEAPON_DEF_BONUS = "def"
Private Const TAG_WEAPON_WIELD_REQ = "weapwieldreq"
Private Const TAG_WEAPON_MIN_BONUS_DMG = "minbonusdmg"

Private Const TAG_WAND_MANAC = "wandmanac"
Private Const TAG_WAND_MAGICD = "wandmagicd"
Private Const TAG_WAND_MELEED = "wandmeleed"
Private Const TAG_WAND_PVM = "wandpvm"
Private Const TAG_WAND_MIN_WIELD_REQ = "wandminwieldreq"
Private Const TAG_WAND_MAX_WIELD_REQ = "wandmaxwieldreq"

Private Const TAG_SALVAGE_MATERIAL = "salvage"    'material ID
Private Const TAG_SALVAGE_COMBINE = "salvagecombine"

Private Const TAG_MIN_WORK = "minwork"
Private Const TAG_MAX_WORK = "maxwork"
Private Const TAG_MAX_VALUE = "maxvalue"

Public Enum eLootFilterType
    FILTER_NONE = 0
    FILTER_ARMOR
    FILTER_WEAPON
    FILTER_WAND
    FILTER_SALVAGE
    NUM_LOOT_FILTERS
End Enum

Public FilterId As Long
Public FilterType As eLootFilterType
Public FilterEnabled As Boolean

'Armor Filter Attribs
Public ArmorType As Long
Public ArmorCoverage As Long
Public ArmorMinAL As Long

'Weapon Filter Attribs
Public WeaponType As Long
Public WeaponMinLowDmg As Integer
Public WeaponMinHighDmg As Integer
Public WeaponElement As Long
Public WeaponMinMod As Integer
Public WeaponMinAttackBonus As Integer
Public WeaponMinDefenseBonus As Integer
Public WeaponMaxWieldReq As Integer
Public WeaponMinBonusDmg As Integer

'Wand Filter Attribs
Public WandManaC As Integer
Public WandMagicD As Integer
Public WandMeleeD As Integer
Public WandPVM As Integer
Public WandMinWieldReq As Integer
Public WandMaxWieldReq As Integer

'Salvage Filter Attribs
Public SalvageMaterial As Long
Public SalvageCombine As Boolean

'Shared Attribs
Public MinWork As Integer
Public MaxWork As Integer
Public MaxValue As Long

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    
    FilterId = 0
    FilterType = FILTER_NONE
    FilterEnabled = True
    MinWork = 1
    MaxWork = 10
    MaxValue = 99999
    
    ArmorType = 0
    ArmorCoverage = 0
    ArmorMinAL = 0
    
    WeaponType = 0
    WeaponMinLowDmg = 0
    WeaponMinHighDmg = 1
    WeaponElement = FL_DMG_ANY
    WeaponMinMod = 0
    WeaponMinAttackBonus = 0
    WeaponMinDefenseBonus = 0
    WeaponMaxWieldReq = 999
    WeaponMinBonusDmg = 0
    
    WandManaC = 0
    WandMagicD = 0
    WandMeleeD = 0
    WandPVM = 0
    WandMinWieldReq = 0
    WandMaxWieldReq = 0
    
    SalvageMaterial = 0
    SalvageCombine = True
    
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               PRIVATE
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Function LoadAttrib(ByRef vAttrib As Variant, ByVal sTag As String, ByVal dat As clsDataEntry) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If dat.ParamExist(sTag) Then
        vAttrib = Val(dat.Param(sTag))
        bRet = True
    End If
    
Fin:
    LoadAttrib = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilter.LoadAttrib(" & sTag & ") - " & Err.Description
    bRet = False
    Resume Fin
End Function

Private Function PassArmorFilter(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not (FilterEnabled) Then
        locDebug "-- Armor Filter #" & FilterId & " is NOT enabled"
        bRet = False
        GoTo Fin
    End If
    
    'Default
    bRet = True
    
    locDebug "--- Armor Filter #" & FilterId & " on " & objItem.Name & "----"
    locDebug "              Filter          Item"
    
    'TODO : Check Armor Type
    If bRet Then bRet = ((ArmorType = ARMORTYPE_ANY) Or (objItem.ArmorType = ArmorType))
    locDebug "ArmorType     " & ArmorType & "          " & objItem.ArmorType & "    Pass:" & CStr(bRet)
    
    'Check Armor Coverage
    If bRet Then bRet = ((ArmorCoverage = ARMORCOVER_ANY) Or (objItem.Coverage = ArmorCoverage))
    locDebug "Coverage     " & ArmorCoverage & "          " & objItem.Coverage & "    Pass:" & CStr(bRet)
    
    'Check Armor Level
    If bRet Then bRet = (objItem.ArmorLevel >= ArmorMinAL)
    locDebug "MinAL     " & ArmorMinAL & "          " & objItem.ArmorLevel & "    Pass:" & CStr(bRet)
    
    'Check Workmanship
    If bRet Then bRet = (objItem.Workmanship <= MaxWork)
    locDebug "MaxWork     " & MaxWork & "          " & objItem.Workmanship & "    Pass:" & CStr(bRet)
    
    'Check Value
    If bRet Then bRet = (objItem.Value <= MaxValue)
    locDebug "MaxValue     " & MaxValue & "          " & objItem.Value & "    Pass:" & CStr(bRet)
    
Fin:
    PassArmorFilter = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLootFilter.PassArmorFilter - " & Err.Description
    Resume Fin
End Function

Private Function PassWeaponFilter(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not (FilterEnabled) Then
        locDebug "-- Weapon Filter #" & FilterId & " is NOT enabled"
        bRet = False
        GoTo Fin
    End If
    
    'Default
    bRet = True
    
    locDebug "--- Weapon Filter #" & FilterId & " on " & objItem.Name & "----"
    locDebug "              Filter          Item"

    If objItem.itemType = ITEM_MISSILE_WEAPON Then
        'Check Missile Weapon Type
        If bRet Then bRet = ((WeaponType = WEAPON_ANY_MISSILE) Or (objItem.SkillUsed = WeaponType))
        locDebug "MissileWeapType     " & WeaponType & "          " & objItem.SkillUsed & "    Pass:" & CStr(bRet)
    
        'Check Damage Mod
        If bRet Then bRet = (objItem.DamageModifier >= WeaponMinMod)
        locDebug "Min Mod%     " & WeaponMinMod & "          " & objItem.DamageModifier & "    Pass:" & CStr(bRet)
        
        'Check Bonus Damage
        If bRet Then bRet = (objItem.ElementBonusDamage >= WeaponMinBonusDmg)
        locDebug "Min Bonus Dmg:  " & WeaponMinBonusDmg & "        " & objItem.ElementBonusDamage & "    Pass:" & CStr(bRet)
    
    Else    'Melee Weapons
        'Check Missile Weapon Type
        If bRet Then bRet = ((WeaponType = WEAPON_ANY_MELEE) Or (objItem.SkillUsed = WeaponType))
        locDebug "MeleeWeapType     " & WeaponType & "          " & objItem.SkillUsed & "    Pass:" & CStr(bRet)
    
        'Check Low/Min Dmg
        If bRet Then bRet = ((objItem.GetLowDamage >= WeaponMinLowDmg) And (objItem.HighDamage >= WeaponMinHighDmg))
        locDebug "MinDmg     " & WeaponMinLowDmg & "-" & WeaponMinHighDmg & "          " & objItem.GetLowDamage & "-" & objItem.HighDamage & "    Pass:" & CStr(bRet)
        
        'Check Attack Bonus
        If bRet Then bRet = (objItem.AttackBonus >= WeaponMinAttackBonus)
        locDebug "Att%     " & WeaponMinAttackBonus & "          " & objItem.AttackBonus & "    Pass:" & CStr(bRet)
        
    End If
    
    'Check Element Type
    If bRet Then bRet = (objItem.DamageFlags = WeaponElement) Or (WeaponElement = FL_DMG_ANY)
    locDebug "Elemt     " & WeaponElement & "          " & objItem.DamageFlags & "    Pass:" & CStr(bRet)
    
    'Check Defense Bonus
    If bRet Then bRet = (objItem.DefenseBonus >= WeaponMinDefenseBonus)
    locDebug "Def%     " & WeaponMinDefenseBonus & "          " & objItem.DefenseBonus & "    Pass:" & CStr(bRet)
    
    'Check Wield Requirement
    If bRet Then bRet = (objItem.WieldReqVal <= WeaponMaxWieldReq)
    locDebug "MaxWieldReq     " & WeaponMaxWieldReq & "          " & objItem.WieldReqVal & "    Pass:" & CStr(bRet)
    
    'Check Workmanship
    If bRet Then bRet = (objItem.Workmanship <= MaxWork)
    locDebug "MaxWork     " & MaxWork & "          " & objItem.Workmanship & "    Pass:" & CStr(bRet)
    
    'Check Value
    If bRet Then bRet = (objItem.Value <= MaxValue)
    locDebug "MaxValue     " & MaxValue & "          " & objItem.Value & "    Pass:" & CStr(bRet)
    
Fin:
    PassWeaponFilter = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLootFilter.PassWeaponFilter - " & Err.Description
    Resume Fin
End Function

Private Function PassWandFilter(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not (FilterEnabled) Then
        locDebug "-- Wand Filter #" & FilterId & " is NOT enabled"
        bRet = False
        GoTo Fin
    End If
    
    'Default
    bRet = True
    
    locDebug "--- Wand Filter #" & FilterId & " on " & objItem.Name & "----"
    locDebug "              Filter          Item"

    'Check ManaC mod
    If bRet Then bRet = (objItem.ManaConvMod >= WandManaC)
    locDebug "WandManaC     " & WandManaC & "          " & objItem.ManaConvMod & "    Pass:" & CStr(bRet)
    
    'Check MagicD mod
    If bRet Then bRet = (objItem.MagicDefense >= WandMagicD)
    locDebug "MagicD%      " & WandMagicD & "          " & objItem.MagicDefense & "    Pass:" & CStr(bRet)
    
    'Check Defense Bonus
    If bRet Then bRet = (objItem.DefenseBonus >= WandMeleeD)
    locDebug "MeleeD%      " & WandMeleeD & "          " & objItem.DefenseBonus & "    Pass:" & CStr(bRet)
    
    'Check PvM Bonus
    If bRet Then bRet = (objItem.PvMBonus >= WandPVM)
    locDebug "PVM          " & WandPVM & "          " & objItem.PvMBonus & "    Pass:" & CStr(bRet)
    
    'Check Wield Requirement
    If bRet Then bRet = (objItem.WieldReqVal <= WandMaxWieldReq)
    locDebug "MaxWieldReq  " & WandMaxWieldReq & "          " & objItem.WieldReqVal & "    Pass:" & CStr(bRet)
    
    'Check Workmanship
    If bRet Then bRet = (objItem.Workmanship <= MaxWork)
    locDebug "MaxWork      " & MaxWork & "          " & objItem.Workmanship & "    Pass:" & CStr(bRet)
    
    'Check Value
    ' JSC - We don't care about Value!!!
    'If bRet Then bRet = (objItem.Value <= MaxValue)
    'locDebug "MaxValue     " & MaxValue & "          " & objItem.Value & "    Pass:" & CStr(bRet)
    
Fin:
    PassWandFilter = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLootFilter.PassWandFilter - " & Err.Description
    Resume Fin
End Function

Private Function PassSalvageFilter(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not (FilterEnabled) Then
        locDebug "-- Salvage Filter #" & FilterId & " is NOT enabled"
        bRet = False
        GoTo Fin
    End If
    
    'Default
    bRet = True
    
    'Check Material Type
    If bRet Then bRet = (objItem.MaterialType = SalvageMaterial)

    'Check Workmanship
    If bRet Then bRet = (objItem.Workmanship >= MinWork)

Fin:
    PassSalvageFilter = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLootFilter.PassSalvageFilter - " & Err.Description
    Resume Fin
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   PUBLIC
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function PassFilter(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    Select Case FilterType
        Case FILTER_ARMOR
            bRet = PassArmorFilter(objItem)
        
        Case FILTER_WEAPON
            bRet = PassWeaponFilter(objItem)
            
        Case FILTER_WAND
            bRet = PassWandFilter(objItem)
            
        Case FILTER_SALVAGE
            bRet = PassSalvageFilter(objItem)
            
        Case Else
            MyDebug "WARNING - clsLootFilter.PassFilter : unsupported FilterType #" & FilterType
            bRet = False
    End Select
    
Fin:
    PassFilter = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLootFilter.PassFilter - " & Err.Description
    Resume Fin
End Function

Public Function AddToDatabase(ByRef db As DataFile) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim dat As New clsDataEntry
    
    dat.AddParam TAG_FILTER_TYPE, FilterType
    dat.AddParam TAG_FILTER_ID, FilterId
    If (FilterEnabled) Then
        dat.AddParam TAG_FILTER_ENABLED, 1
    Else
        dat.AddParam TAG_FILTER_ENABLED, 0
    End If
    
    Select Case FilterType
        Case FILTER_ARMOR
            dat.AddParam TAG_ARMOR_TYPE, ArmorType
            dat.AddParam TAG_ARMOR_COVERAGE, ArmorCoverage
            dat.AddParam TAG_ARMOR_MIN_AL, ArmorMinAL
            dat.AddParam TAG_MAX_WORK, MaxWork
            dat.AddParam TAG_MAX_VALUE, MaxValue
            
        Case FILTER_WEAPON
            dat.AddParam TAG_WEAPON_TYPE, WeaponType
            If g_ACConst.WeaponSkills.Item(WeaponType).Flag = ITEM_MELEE_WEAPON Then 'Melee weapons
                dat.AddParam TAG_WEAPON_ATT_BONUS, WeaponMinAttackBonus
                dat.AddParam TAG_WEAPON_MIN_LOW_DMG, WeaponMinLowDmg
                dat.AddParam TAG_WEAPON_MIN_HIGH_DMG, WeaponMinHighDmg
            Else    'Missile Weapons
                dat.AddParam TAG_WEAPON_MIN_DMG_MOD, WeaponMinMod
                dat.AddParam TAG_WEAPON_MIN_BONUS_DMG, WeaponMinBonusDmg
            End If
            dat.AddParam TAG_WEAPON_ELEMENT, WeaponElement
            dat.AddParam TAG_WEAPON_DEF_BONUS, WeaponMinDefenseBonus
            dat.AddParam TAG_WEAPON_WIELD_REQ, WeaponMaxWieldReq
            dat.AddParam TAG_MAX_WORK, MaxWork
            dat.AddParam TAG_MAX_VALUE, MaxValue
            
        Case FILTER_WAND
            dat.AddParam TAG_WAND_MANAC, WandManaC
            dat.AddParam TAG_WAND_MAGICD, WandMagicD
            dat.AddParam TAG_WAND_MELEED, WandMeleeD
            dat.AddParam TAG_WAND_PVM, WandPVM
            dat.AddParam TAG_WAND_MIN_WIELD_REQ, WandMinWieldReq
            dat.AddParam TAG_WAND_MAX_WIELD_REQ, WandMaxWieldReq
            dat.AddParam TAG_MAX_WORK, MaxWork
            dat.AddParam TAG_MAX_VALUE, MaxValue
        
        Case FILTER_SALVAGE
            dat.AddParam TAG_SALVAGE_MATERIAL, SalvageMaterial
            dat.AddParam TAG_MIN_WORK, MinWork
            'SalvageCombine
            If (SalvageCombine) Then
                dat.AddParam TAG_SALVAGE_COMBINE, 1
            Else
                dat.AddParam TAG_SALVAGE_COMBINE, 0
            End If
    End Select
    
    bRet = db.AddData(dat)
    
Fin:
    AddToDatabase = bRet
    Set dat = Nothing
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilter.SaveToDatabase - " & Err.Description
    bRet = False
    Resume Fin
End Function


Public Function LoadFromDataEntry(ByVal dat As clsDataEntry) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim isEnabled As Integer
    Dim salvCombine As Integer
    
    If LoadAttrib(FilterType, TAG_FILTER_TYPE, dat) Then
        LoadAttrib FilterId, TAG_FILTER_ID, dat
        LoadAttrib isEnabled, TAG_FILTER_ENABLED, dat
        If (isEnabled <> 0) Then
            FilterEnabled = True
        Else
            FilterEnabled = False
        End If
        
        LoadAttrib ArmorType, TAG_ARMOR_TYPE, dat
        LoadAttrib ArmorCoverage, TAG_ARMOR_COVERAGE, dat
        LoadAttrib ArmorMinAL, TAG_ARMOR_MIN_AL, dat
        
        LoadAttrib WeaponType, TAG_WEAPON_TYPE, dat
        LoadAttrib WeaponMinLowDmg, TAG_WEAPON_MIN_LOW_DMG, dat
        LoadAttrib WeaponMinHighDmg, TAG_WEAPON_MIN_HIGH_DMG, dat
        LoadAttrib WeaponElement, TAG_WEAPON_ELEMENT, dat
        LoadAttrib WeaponMinMod, TAG_WEAPON_MIN_DMG_MOD, dat
        LoadAttrib WeaponMinAttackBonus, TAG_WEAPON_ATT_BONUS, dat
        LoadAttrib WeaponMinDefenseBonus, TAG_WEAPON_DEF_BONUS, dat
        LoadAttrib WeaponMaxWieldReq, TAG_WEAPON_WIELD_REQ, dat
        LoadAttrib WeaponMinBonusDmg, TAG_WEAPON_MIN_BONUS_DMG, dat
        
        LoadAttrib WandManaC, TAG_WAND_MANAC, dat
        LoadAttrib WandMagicD, TAG_WAND_MAGICD, dat
        LoadAttrib WandMeleeD, TAG_WAND_MELEED, dat
        LoadAttrib WandPVM, TAG_WAND_PVM, dat
        LoadAttrib WandMinWieldReq, TAG_WAND_MIN_WIELD_REQ, dat
        LoadAttrib WandMaxWieldReq, TAG_WAND_MAX_WIELD_REQ, dat
        
        LoadAttrib SalvageMaterial, TAG_SALVAGE_MATERIAL, dat
        
        LoadAttrib MinWork, TAG_MIN_WORK, dat
        LoadAttrib MaxWork, TAG_MAX_WORK, dat
        LoadAttrib MaxValue, TAG_MAX_VALUE, dat

        LoadAttrib salvCombine, TAG_SALVAGE_COMBINE, dat
        If (salvCombine <> 0) Then
            SalvageCombine = True
        Else
            SalvageCombine = False
        End If
    
        bRet = True
    End If
    
Fin:
    LoadFromDataEntry = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilter.LoadFromDataEntry - " & Err.Description
    bRet = False
    Resume Fin
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                       UTILS
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub locDebug(ByVal sDebugMsg As String, Optional ByVal bSilent As Boolean = False)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[LootFilter #" & FilterId & "] " & sDebugMsg, bSilent)
    End If
End Sub
