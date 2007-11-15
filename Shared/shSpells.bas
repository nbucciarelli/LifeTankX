Attribute VB_Name = "shSpells"
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                 SHARED MODULE                       [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                                                     [[
'[[                     Spells                          [[
'[[                                                     [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Option Explicit

Public Const TAG_SPELL_UNKNOWN = "Unknown"
Public Const TAG_SPELL_NA = "n/a"

Public Const TAG_SPELL_NAME = "spl"
Public Const TAG_SPELL_TYPE = "type"
Public Const TAG_SPELL_FAMILY = "family"
Public Const TAG_SPELL_SCHOOL = "sch"
Public Const TAG_SPELL_LEVEL = "lvl"
Public Const TAG_SPELL_ELEMENT = "elemt"
Public Const TAG_SPELL_ID = "id"
Public Const TAG_SPELL_ICON = "icon"
Public Const TAG_SPELL_DESCRIPTION = "desc"

Public Const NO_FAMILY = "NoFamily"

'-----------------------------
' Magic Schools
'-----------------------------

Public Enum eMagicSchools
    SCHOOL_LIFE
    SCHOOL_WAR
    SCHOOL_CREATURE
    SCHOOL_ITEM
    SCHOOL_UNKNOWN
End Enum

Public Const STR_SCHOOL_LIFE = "Life"
Public Const STR_SCHOOL_WAR = "War"
Public Const STR_SCHOOL_CREATURE = "Creature"
Public Const STR_SCHOOL_ITEM = "Item"
Public Const STR_SCHOOL_UNKNOWN = "Unknown Magic School"

'-----------------------------
' Damage Types
'-----------------------------

Public Enum eDamageType
    DMG_SLASHING = 0
    DMG_BLUDGEONING
    DMG_PIERCING
    DMG_FIRE
    DMG_COLD
    DMG_ACID
    DMG_LIGHTNING
    DMG_NONE
End Enum

Public Const STR_DMG_SLASHING = "Slash"
Public Const STR_DMG_BLUDGEONING = "Bludgeon"
Public Const STR_DMG_PIERCING = "Pierce"
Public Const STR_DMG_FIRE = "Fire"
Public Const STR_DMG_COLD = "Cold"
Public Const STR_DMG_ACID = "Acid"
Public Const STR_DMG_LIGHTNING = "Lightning"
Public Const STR_DMG_NONE = "None"


'-----------------------------
' Spell Type
'-----------------------------
Public Enum eSpellType
    SPELLTYPE_NORMAL
    SPELLTYPE_BOLT
    SPELLTYPE_ARC
    SPELLTYPE_VOLLEY
    SPELLTYPE_WAVE
    SPELLTYPE_WALL
    SPELLTYPE_RING
    SPELLTYPE_STREAK
    SPELLTYPE_VULN
    SPELLTYPE_LURE
    SPELLTYPE_LIFEPRO
    SPELLTYPE_BANE
    SPELLTYPE_BLAST
    SPELLTYPE_TRANSPORT
    
    NUM_SPELLTYPES
End Enum

Public Const STR_SPELLTYPE_NORMAL = "Normal"
Public Const STR_SPELLTYPE_BOLT = "Bolt"
Public Const STR_SPELLTYPE_ARC = "Arc"
Public Const STR_SPELLTYPE_VOLLEY = "Volley"
Public Const STR_SPELLTYPE_WAVE = "Wave"
Public Const STR_SPELLTYPE_WALL = "Wall"
Public Const STR_SPELLTYPE_RING = "Ring"
Public Const STR_SPELLTYPE_STREAK = "Streak"
Public Const STR_SPELLTYPE_VULN = "Vuln"
Public Const STR_SPELLTYPE_LURE = "Lure"
Public Const STR_SPELLTYPE_LIFEPRO = "Pro"
Public Const STR_SPELLTYPE_BANE = "Bane"
Public Const STR_SPELLTYPE_BLAST = "Blast"
Public Const STR_SPELLTYPE_TRANSPORT = "Transport"


'**************************************************************************

Public Function GetSchoolString(ByVal iSchoolId) As String
Dim sRet As String
    Select Case iSchoolId
        Case SCHOOL_LIFE
            sRet = STR_SCHOOL_LIFE
        Case SCHOOL_WAR
            sRet = STR_SCHOOL_WAR
        Case SCHOOL_CREATURE
            sRet = STR_SCHOOL_CREATURE
        Case SCHOOL_ITEM
            sRet = STR_SCHOOL_ITEM
        Case Else
            sRet = STR_SCHOOL_UNKNOWN
    End Select
    
    GetSchoolString = sRet
End Function

Public Function GetSchoolId(ByVal sSchoolName) As Integer
Dim iRet As Integer

    Select Case sSchoolName
        Case STR_SCHOOL_LIFE
            iRet = SCHOOL_LIFE
        Case STR_SCHOOL_WAR
            iRet = SCHOOL_WAR
        Case STR_SCHOOL_CREATURE
            iRet = SCHOOL_CREATURE
        Case STR_SCHOOL_ITEM
            iRet = SCHOOL_ITEM
        Case Else
            iRet = SCHOOL_UNKNOWN
    End Select
    
    GetSchoolId = iRet
End Function

'**************************************************************************

Public Function GetSpelltypeString(ByVal iTypeId As Integer) As String
Dim sRet As String
    Select Case iTypeId
        Case SPELLTYPE_BOLT
            sRet = STR_SPELLTYPE_BOLT
        Case SPELLTYPE_ARC
            sRet = STR_SPELLTYPE_ARC
        Case SPELLTYPE_VOLLEY
            sRet = STR_SPELLTYPE_VOLLEY
        Case SPELLTYPE_WAVE
            sRet = STR_SPELLTYPE_WAVE
        Case SPELLTYPE_WALL
            sRet = STR_SPELLTYPE_WALL
        Case SPELLTYPE_RING
            sRet = STR_SPELLTYPE_RING
        Case SPELLTYPE_STREAK
            sRet = STR_SPELLTYPE_STREAK
        Case SPELLTYPE_VULN
            sRet = STR_SPELLTYPE_VULN
        Case SPELLTYPE_LURE
            sRet = STR_SPELLTYPE_LURE
        Case SPELLTYPE_LIFEPRO
            sRet = STR_SPELLTYPE_LIFEPRO
        Case SPELLTYPE_BANE
            sRet = STR_SPELLTYPE_BANE
        Case SPELLTYPE_BLAST
            sRet = STR_SPELLTYPE_BLAST
        Case SPELLTYPE_TRANSPORT
            sRet = STR_SPELLTYPE_TRANSPORT
        Case Else
            sRet = STR_SPELLTYPE_NORMAL
    End Select
    
    GetSpelltypeString = sRet
End Function

Public Function GetSpelltypeId(ByVal sTypeString As String) As Integer
Dim iRet As Integer

    Select Case sTypeString
        Case STR_SPELLTYPE_BOLT
            iRet = SPELLTYPE_BOLT
        Case STR_SPELLTYPE_ARC
            iRet = SPELLTYPE_ARC
        Case STR_SPELLTYPE_VOLLEY
            iRet = SPELLTYPE_VOLLEY
        Case STR_SPELLTYPE_WAVE
            iRet = SPELLTYPE_WAVE
        Case STR_SPELLTYPE_WALL
            iRet = SPELLTYPE_WALL
        Case STR_SPELLTYPE_RING
            iRet = SPELLTYPE_RING
        Case STR_SPELLTYPE_STREAK
            iRet = SPELLTYPE_STREAK
        Case STR_SPELLTYPE_VULN
            iRet = SPELLTYPE_VULN
        Case STR_SPELLTYPE_LURE
            iRet = SPELLTYPE_LURE
        Case STR_SPELLTYPE_LIFEPRO
            iRet = SPELLTYPE_LIFEPRO
        Case STR_SPELLTYPE_BANE
            iRet = SPELLTYPE_BANE
        Case STR_SPELLTYPE_BLAST
            iRet = SPELLTYPE_BLAST
        Case STR_SPELLTYPE_TRANSPORT
            iRet = SPELLTYPE_TRANSPORT
        Case Else
            iRet = SPELLTYPE_NORMAL
    End Select
    
    GetSpelltypeId = iRet
End Function

'**************************************************************************

Public Function GetDamageString(ByVal dmgType As Integer) As String
    Select Case dmgType
        Case DMG_SLASHING
            GetDamageString = STR_DMG_SLASHING
        Case DMG_BLUDGEONING
            GetDamageString = STR_DMG_BLUDGEONING
        Case DMG_PIERCING
            GetDamageString = STR_DMG_PIERCING
        Case DMG_FIRE
            GetDamageString = STR_DMG_FIRE
        Case DMG_COLD
            GetDamageString = STR_DMG_COLD
        Case DMG_ACID
            GetDamageString = STR_DMG_ACID
        Case DMG_LIGHTNING
            GetDamageString = STR_DMG_LIGHTNING
        Case Else
            GetDamageString = STR_DMG_NONE
    End Select
End Function

Public Function GetDamageType(ByVal dmgName As String) As Integer
    Select Case dmgName
        Case STR_DMG_SLASHING
            GetDamageType = DMG_SLASHING
        Case STR_DMG_BLUDGEONING
            GetDamageType = DMG_BLUDGEONING
        Case STR_DMG_PIERCING
            GetDamageType = DMG_PIERCING
        Case STR_DMG_FIRE
            GetDamageType = DMG_FIRE
        Case STR_DMG_COLD
            GetDamageType = DMG_COLD
        Case STR_DMG_ACID
            GetDamageType = DMG_ACID
        Case STR_DMG_LIGHTNING
            GetDamageType = DMG_LIGHTNING
        Case Else
            GetDamageType = DMG_NONE
    End Select
End Function
