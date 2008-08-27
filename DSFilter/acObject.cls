VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const VULN_EXPIRATION_TIME = 240 '4 minutes

'---------------------------------------------------------
' Enums
'---------------------------------------------------------
Public Enum eObjectType
    TYPE_UNKNOWN
    TYPE_PLAYER
    TYPE_MONSTER
    TYPE_MERCHANT
    TYPE_ITEM
    TYPE_JUNK
End Enum

Public Enum ePlayerType
    PLAYER_WHITE
    PLAYER_PINK
    PLAYER_RED
End Enum

'---------------------------------------------------------
' Common Attributes
'---------------------------------------------------------
Private m_UserData() As Variant         'Resizable Array for extra user data

Public Name As String                   'object name as it appears ingame
Public GUID As Long                     'object identifer
Public Icon As Integer
Public ObjectType As eObjectType
Public PlayerType As ePlayerType
Public ItemType As eItemTypes
Public Loc As acLoc                     'object location
Public timeData As Long                 'When to delete
Public canDelete As Boolean             'Can we delete

'---------------------------------------------------------
' Attributes Available @ Object Creation (or after IDing)
'---------------------------------------------------------
Public Wielder As Long              'GUID of the object who's currently wearing/equipping this item
Public Container As Long            'GUID of the object holding/carrying this item

Public Coverage As Long
Public Coverage2 As Long
Public Coverage3 As Long            'the flag we're interested in

Public UseType As Long
Public Value As Long
Public Burden As Long
Public Workmanship As Long
Public MaterialType As Long
Public UsesLeft As Integer
Public TotalUses As Integer
Public StackCount As Integer
Public StackMax As Integer
Public AssociatedSpellId As Integer

Public MonarchID As Long            'the player's monarch id

'---------------------------------------------------------
' Attributes Available ONLY AFTER IDing the object
'---------------------------------------------------------
Public LastIDTime As Double             'Last time this item has been IDed
Public KillerName As String             'Name of the entity who got kill on this object
Public Inscription As String
Public Inscriber As String
Public Description As String
Public ShortDesc As String
Public UsageInstructions As String

Public TinkCount As Integer
Public TinkerName As String
Public Imbue As String
Public LoreReq As Integer
Public ActivateSkill As String
Public ActivateSkillVal As Integer
Public RaceReq As String
Public RankReq As Integer

'Armor Info
Public ArmorLevel As Integer
Public ArmorType As Long

'Weapon Info
Public WieldReqType As Long
Public WieldReqVal As Long
Public WieldReqId As Long
Public SkillReqId As Long

Public ElementBonusDamage As Long       'Elemental Bonus Damage
Public DamageFlags As Long              'Element of the weapon (eWeaponDamageFlags combinations)
Public DamageType As Integer
Public SkillUsed As Integer
Public HighDamage As Long
Public DamageModifier As Integer        'X, where X is the +X% damage modifier (+120%, +130%...)
Public Variance As Double               'Weapon damage variance
Public DefenseBonus As Integer
Public AttackBonus As Integer
Public MagicDefense As Single
Public MissileDefense As Integer
Public BitingStrike As Boolean
Public CrushingBlow As Boolean
Public ResistanceCleaving As String
Public ResistanceCleavingType As String
Public slayerType As Integer

'Wand info
Public ManaConvMod As Integer
Public PvMBonus As Integer

'---------------------------------------------------------
' Stats Info [ after ID ]
'---------------------------------------------------------
Public Health As Integer
Public MaxHealth As Integer
Public Stamina As Integer
Public MaxStamina As Integer
Public Mana As Integer
Public MaxMana As Integer

Public AttribStrenght As Integer
Public AttribEndurance As Integer
Public AttribCoordination As Integer
Public AttribQuickness As Integer
Public AttribFocus As Integer
Public AttribSelf As Integer

'---------------------------------------------------------
' Player Info [ after ID ]
'---------------------------------------------------------
Public MonarchName As String            'not working
Public FellowshipName As String         'name of the fellowship this player is in
Public Rank As Integer
Public Gender As String
Public Followers As Integer
Public Leadership As Integer
Public race As String
Public Class As String
Public Level As Integer                 'level of this character

'---------------------------------------------------------
' Extra Object Data
'---------------------------------------------------------
Public Equiped As Boolean               'is item worn by our character ?
Public Dead As Boolean                  'Has this object been killed?
Public IsPack As Boolean                'is it a pack/foci in our inventory?
Public ShareLoot As Boolean             'if the fellow player has Share Loot on
Public HasMinors As Boolean             'Indicates that this item holds at least 1 minor spell
Public HasMajors As Boolean             'Indicates that this item holds at least 1 major spell
Public HasEpics As Boolean              'Indicates that this item holds at least 1 epic spell
Public IsRare As Boolean                'Indicates that this item is a Rare
Public RareNumber As Long               'Number of the rare item.
Public Spells As Dictionary             'List of inate spells on this item
Public SpellsActive As Dictionary       'List of Active spells on this item
Public Spellcraft As Long               'Spellcraft of the item.
Public unEnchantable As Boolean         'Is this item unEnchantable?

'---------------------------------------------------------
' Vulns Info
'---------------------------------------------------------
'Public Vulns As Integer                 'vuln flags
'Public Imperiled As Boolean
'Public Yielded As Boolean
Public Vulnerability As Integer

Private m_Vulns As Integer 'vuln flags
Private m_fVulnTimer(DMG_SLASHING To DMG_LIGHTNING) As Double

Private m_bImperilOrBludgeon As Boolean
    Private m_fImperilOrBludgeonTimer As Double

Private m_bImperiled As Boolean
    Private m_fImperilTimer As Double
Private m_bYielded As Boolean
    Private m_fYieldTimer As Double
    
'Extra...
Public gameDataType As Long
Public Priority As Integer
Public VendorMaxBuy As Long
Public VendorFractBuy As Single
Public VendorFractSell As Single


'#########################################################
'#            Constructor / Destructor
'#########################################################

Private Sub Class_Initialize()
    Name = "UnknownEntity"
    GUID = 0
    Icon = 0
    ObjectType = TYPE_UNKNOWN
    PlayerType = PLAYER_WHITE
    ItemType = ITEM_UNKNOWN
    Set Loc = New acLoc
    timeData = 0
    canDelete = False
    
    Call SetUserDataCount(0)
    
    Wielder = 0
    Container = 0
    UseType = 0
    Coverage = 0
    Coverage2 = 0
    Coverage3 = 0
    Value = 0
    Burden = 0
    Workmanship = 0
    MaterialType = 0
    UsesLeft = 0
    TotalUses = 0
    StackCount = 1
    StackMax = 1
    MonarchID = 0
    KillerName = "n/a"
    Description = "n/a"
    ShortDesc = "n/a"
    UsageInstructions = "n/a"
    Inscription = "n/a"
    Inscriber = "n/a"
    TinkCount = 0
    TinkerName = ""
    Imbue = ""
    ArmorLevel = 0
    ArmorType = -1 'ARMORTYPE_UNKNOWN
    LoreReq = 0
    WieldReqType = 0
    WieldReqVal = 0
    WieldReqId = 0
    SkillReqId = 0
    RaceReq = ""
    RankReq = 0
    ElementBonusDamage = 0
    DamageFlags = 0
    DamageType = DMG_NONE
    SkillUsed = 0
    HighDamage = 0
    Variance = 0
    DamageModifier = 0
    DefenseBonus = 0
    AttackBonus = 0
    ManaConvMod = 0
    PvMBonus = 0
    AttribStrenght = -1
    AttribEndurance = -1
    AttribCoordination = -1
    AttribQuickness = -1
    AttribFocus = -1
    AttribSelf = -1
    Health = -1
    MaxHealth = -1
    Stamina = -1
    MaxStamina = -1
    Mana = -1
    MaxMana = -1
    MonarchName = "n/a"
    FellowshipName = "n/a"
    Rank = -1
    Gender = "n/a"
    Followers = -1
    Leadership = -1
    race = "n/a"
    Class = "n/a"
    Level = -1
    Equiped = False
    Dead = False
    IsPack = False
    ShareLoot = False
    HasMinors = False
    HasMajors = False
    IsRare = False
    RareNumber = 0
    AssociatedSpellId = 0
    Priority = 0
    Spellcraft = 0
    unEnchantable = False
    
    VendorMaxBuy = 0
    VendorFractBuy = 0
    VendorFractSell = 0
    
    MagicDefense = 0
    MissileDefense = 0
    
    BitingStrike = False
    CrushingBlow = False
    
    ResistanceCleaving = ""
    ResistanceCleavingType = ""
    
    slayerType = 0

    Vulnerability = -1 'DMG_SLASHING
    m_Vulns = -1
    m_bImperiled = False
    m_bYielded = False
    
    m_bImperilOrBludgeon = False
    m_fImperilOrBludgeonTimer = 0
    LastIDTime = 0
    Set Spells = New Dictionary
    Set SpellsActive = New Dictionary
    
End Sub

Private Sub Class_Terminate()
    Set Loc = Nothing
    Set Spells = Nothing
    Set SpellsActive = Nothing
End Sub

'#########################################################
'#                      Properties
'#########################################################

Public Property Get UserDataCount() As Long
    UserDataCount = UBound(m_UserData) + 1
End Property

Public Property Get UserData(ByVal iIndex As Integer) As Variant
On Error GoTo ErrorHandler

    If iIndex >= 0 And iIndex < UserDataCount Then
        UserData = m_UserData(iIndex)
    Else
        UserData = 0
    End If
    
Fin:
    Exit Property
ErrorHandler:
    myError "acObject.UserData(" & iIndex & ") - " & Err.Description
    Resume Fin
End Property


Public Function SetUserData(ByVal iIndex As Integer, ByVal vValue As Variant) As Boolean
On Error GoTo ErrorHandler

    If iIndex >= 0 And iIndex < UserDataCount Then
        m_UserData(iIndex) = vValue
        SetUserData = True
    Else
        SetUserData = False
    End If

Fin:
    Exit Function
ErrorHandler:
    myError "acObject.SetUserData(" & iIndex & ") - " & Err.Description
    Resume Fin
End Function

'#########################################################
'#                   Public Methods
'#########################################################

Public Function IsOnGround() As Boolean
    IsOnGround = ((Wielder = 0) And (Container = 0))
End Function

Public Function IsWielded() As Boolean
    IsWielded = ((Wielder <> 0) And (Container = 0))
End Function

Public Function GetLowDamage() As Double
    GetLowDamage = Round(CDbl(HighDamage) - Variance * CDbl(HighDamage), 2)
End Function

Public Function GetSquareRange() As Single
On Error GoTo ErrorHandler

    If Not Valid(g_Objects.Player) Then
        myDebug "acObject.GetSquareRange: Player is null. Returning -1"
        GetSquareRange = -1
        Exit Function
    End If

    If GUID = g_Objects.Player.GUID Then
        GetSquareRange = 0
        Exit Function
    End If

    GetSquareRange = g_Objects.Player.Loc.SquareDistanceTo(Loc)
    
Fin:
    Exit Function
ErrorHandler:
    GetSquareRange = -1
    myError "acObject.GetSquareRange - " & Err.Description
    Resume Fin
End Function

Public Function GetRange() As Single
On Error GoTo ErrorHandler

    GetRange = Sqr(GetSquareRange)
    
Fin:
    Exit Function
ErrorHandler:
    GetRange = -1
    myError "acObject.GetRange - " & Err.Description
    Resume Fin
End Function

Public Function SetUserDataCount(ByVal iCount As Integer) As Boolean
On Error GoTo ErrorHandler

    If iCount >= 0 Then
        ReDim Preserve m_UserData(0 To iCount)
        SetUserDataCount = True
    Else
        SetUserDataCount = False
    End If
    
Fin:
    Exit Function
ErrorHandler:
    SetUserDataCount = False
    myError "acObject.SetUserDataCount(" & iCount & ") - " & Err.Description
    Resume Fin
End Function

Public Function Clone() As acObject
On Error GoTo ErrorHandler
    
    Dim objClone As New acObject
    
    With objClone
        .Name = Name
        .GUID = GUID
        .Icon = Icon
        .Wielder = Wielder
        .Container = Container
        .ObjectType = ObjectType
        .PlayerType = PlayerType
        .ItemType = ItemType
        .UseType = UseType
        .Dead = Dead
        .timeData = timeData
        .canDelete = canDelete
        .Coverage = Coverage
        .Coverage2 = Coverage2
        .Coverage3 = Coverage3
        .Value = Value
        .Burden = Burden
        .Workmanship = Workmanship
        .MaterialType = MaterialType
        .UsesLeft = UsesLeft
        .TotalUses = TotalUses
        .StackCount = StackCount
        .StackMax = StackMax
        .MonarchID = MonarchID
        .KillerName = KillerName
        .Description = Description
        .ShortDesc = ShortDesc
        .UsageInstructions = UsageInstructions
        .Inscription = Inscription
        .Inscriber = Inscriber
        .TinkCount = TinkCount
        .TinkerName = TinkerName
        .Imbue = Imbue
        .ArmorLevel = ArmorLevel
        .ArmorType = ArmorType
        .LoreReq = LoreReq
        .WieldReqType = WieldReqType
        .WieldReqVal = WieldReqVal
        .WieldReqId = WieldReqId
        .SkillReqId = SkillReqId
        .RaceReq = RaceReq
        .RankReq = RankReq
        .DamageFlags = DamageFlags
        .DamageType = DamageType
        .SkillUsed = SkillUsed
        .HighDamage = HighDamage
        .DamageModifier = DamageModifier
        .Variance = Variance
        .ElementBonusDamage = ElementBonusDamage
        .DefenseBonus = DefenseBonus
        .AttackBonus = AttackBonus
        .ManaConvMod = ManaConvMod
        .PvMBonus = PvMBonus
        .AttribStrenght = AttribStrenght
        .AttribEndurance = AttribEndurance
        .AttribCoordination = AttribCoordination
        .AttribQuickness = AttribQuickness
        .AttribFocus = AttribFocus
        .AttribSelf = AttribSelf
        .Health = Health
        .MaxHealth = MaxHealth
        .Stamina = Stamina
        .MaxStamina = MaxStamina
        .Mana = Mana
        .MaxMana = MaxMana
        .MonarchName = MonarchName
        .FellowshipName = FellowshipName
        .Rank = Rank
        .Gender = Gender
        .Followers = Followers
        .Leadership = Leadership
        .race = race
        .Class = Class
        .Level = Level
        .Equiped = Equiped
        .IsPack = IsPack
        .ShareLoot = ShareLoot
        .HasMinors = HasMinors
        .HasMajors = HasMajors
        .IsRare = IsRare
        .RareNumber = RareNumber
        .Imperiled = Imperiled
        .Vulnerability = Vulnerability
        '.Vulns = Vulns
        .Spellcraft = Spellcraft
        .AssociatedSpellId = AssociatedSpellId
        
        .MagicDefense = MagicDefense
        .MissileDefense = MissileDefense
        
        .BitingStrike = BitingStrike
        .CrushingBlow = CrushingBlow
        
        .slayerType = slayerType
        
        .ResistanceCleaving = ResistanceCleaving
        .ResistanceCleavingType = ResistanceCleavingType
        
        .VendorMaxBuy = VendorMaxBuy
        .VendorFractSell = VendorFractSell
        .VendorFractBuy = VendorFractBuy

        'Clone location
        Set .Loc = Loc.Clone
        
        'Clone User Data
        If .SetUserDataCount(Me.UserDataCount) Then
            Dim i As Integer
            For i = 0 To .UserDataCount - 1
                .UserData(i) = Me.UserData(i)
            Next i
        End If
        
    End With
    
Fin:
    Set Clone = objClone
    Set objClone = Nothing
    Exit Function
ErrorHandler:
    Set objClone = Nothing
    myError "acObject.Clone - " & Err.Description
    Resume Fin
End Function

Public Function GetItemTypeName() As String
    GetItemTypeName = g_Const.GetItemTypeName(ItemType)
End Function

Public Sub FixImbueString()
On Error GoTo ErrorHandler
myError "Fixing Imbue String..."
If InStr(Imbue, "Crushing Blow, Biting StrikeCrushing Blow, Biting Strike") > 0 Then
    myError "Found a duplicate imbue. Replacing..."
    Imbue = Replace(Imbue, "Crushing Blow, Biting StrikeCrushing Blow, Biting Strike", "Crushing Blow, Biting Strike,")
    Imbue = Replace(Imbue, ",,", "")
    Imbue = Replace(Imbue, ",,", "")
End If

Fin:
    Exit Sub
ErrorHandler:
    myError "acObjects.FixImbueString - " & Err.Description
    Resume Fin
End Sub
Public Sub UpdateArmorType()
On Error GoTo ErrorHandler

    Dim oData As DataItem
    
    For Each oData In g_Const.Armors
        If InStr(1, LCase(Name), LCase(oData.Val)) Then
            ArmorType = oData.Key
            GoTo Fin
        End If
    Next oData
    
    ArmorType = ARMORTYPE_UNKNOWN
    
Fin:
    Exit Sub
ErrorHandler:
    myError "acObjects.UpdateArmorType - " & Err.Description
    Resume Fin
End Sub

'----------------------------
'Debuffs
'----------------------------

Public Property Get ImpOrBludgeon() As Boolean
    If m_bImperilOrBludgeon And TimerExpired(m_fImperilOrBludgeonTimer) Then
        m_bImperilOrBludgeon = False
    End If
    ImpOrBludgeon = m_bImperilOrBludgeon
End Property

Public Property Let ImpOrBludgeon(ByVal bVal As Boolean)
    If bVal Then
        m_fImperilOrBludgeonTimer = g_Time + VULN_EXPIRATION_TIME
    Else
        m_fImperilOrBludgeonTimer = 0
    End If
    m_bImperilOrBludgeon = bVal
End Property

Public Property Get Imperiled() As Boolean
    If m_bImperiled And TimerExpired(m_fImperilTimer) Then
        m_bImperiled = False
    End If
    Imperiled = m_bImperiled
End Property

Public Property Let Imperiled(ByVal bVal As Boolean)
    If bVal Then
        m_fImperilTimer = g_Time + VULN_EXPIRATION_TIME
    Else
        m_fImperilTimer = 0
    End If
    m_bImperiled = bVal
End Property

Public Function ImperilTimeRemaining() As Double
    If Not m_bImperiled Or TimerExpired(m_fImperilTimer) Then
        ImperilTimeRemaining = 0
    Else
        ImperilTimeRemaining = TimeRemaining(m_fImperilTimer)
    End If
End Function

Public Property Get Yielded() As Boolean
    If m_bYielded And TimerExpired(m_fYieldTimer) Then
        m_bYielded = False
    End If
    Yielded = m_bYielded
End Property

Public Property Let Yielded(ByVal bVal As Boolean)
    If bVal Then
        m_fYieldTimer = g_Time + VULN_EXPIRATION_TIME
    Else
        m_fYieldTimer = 0
    End If
    m_bYielded = bVal
End Property

Public Function YieldTimeRemaining() As Double
    If Not m_bYielded Or TimerExpired(m_fYieldTimer) Then
        YieldTimeRemaining = 0
    Else
        YieldTimeRemaining = TimeRemaining(m_fYieldTimer)
    End If
End Function

Private Function VulnExpired(ByVal iDmg As Integer) As Boolean
On Error GoTo ErrorHandler

    VulnExpired = TimerExpired(m_fVulnTimer(iDmg))
    
Fin:
    Exit Function
ErrorHandler:
    VulnExpired = False
    myError "acObject.VulnExpired - " & Err.Description
    Resume Fin
End Function

Private Sub UpdateVulns()
On Error GoTo ErrorHandler

    Dim i As Integer

    'myError "UpdateVulns"
    For i = DMG_SLASHING To DMG_LIGHTNING
        If IsVulnedTo(i) And VulnExpired(i) Then
            Call RemoveVuln(i)
        End If
    Next i
   
Fin:
    Exit Sub
ErrorHandler:
    myError "acObject.UpdateVulns - " & Err.Description
    Resume Fin
End Sub

Public Property Get Vulns() As Integer
    Call UpdateVulns
    Vulns = m_Vulns
End Property

Public Sub AddVuln(ByVal iDamageType As Integer)
    myDebug Name & " : adding vuln " & iDamageType
    m_Vulns = m_Vulns Or dmg2vuln(iDamageType)
    m_fVulnTimer(iDamageType) = g_Time + VULN_EXPIRATION_TIME
End Sub

Public Sub RemoveVuln(ByVal iDamageType As Integer)
On Error GoTo ErrorHandler

    myDebug Name & " : removing vuln " & iDamageType
    m_Vulns = m_Vulns And Not (dmg2vuln(iDamageType))
    m_fVulnTimer(iDamageType) = 0
    
    Call g_Objects.RaiseOnVulnExpired(Me)

Fin:
    Exit Sub
ErrorHandler:
    myError "acObject.RemoveVuln - " & Err.Description
    Resume Fin
End Sub

Public Sub RemoveAllDebuffs()
    Dim i As Integer
    
    m_Vulns = 0
    For i = DMG_SLASHING To DMG_LIGHTNING
        m_fVulnTimer(i) = 0
    Next i
    
    m_fImperilTimer = 0
    m_bImperiled = False
    
    m_fYieldTimer = 0
    m_bYielded = False
    
End Sub

'returns time in seconds
Public Function VulnTimeRemaining(ByVal iDamageType As Integer) As Double
    If Not IsVulnedTo(iDamageType) Or VulnExpired(iDamageType) Then
        VulnTimeRemaining = 0
    Else
        VulnTimeRemaining = TimeRemaining(m_fVulnTimer(iDamageType))
    End If
End Function

Public Function IsVulnedTo(ByVal DamageType As Integer) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    
    bRet = (m_Vulns And dmg2vuln(DamageType))
    
    If bRet And VulnExpired(DamageType) Then
        Call RemoveVuln(DamageType)
        bRet = False
    End If
    
Fin:
    IsVulnedTo = bRet
    Exit Function
ErrorHandler:
    bRet = False
    myError "acObject.IsVulnedTo - " & Err.Description
    Resume Fin
End Function

Public Function CountVulns() As Integer
    Dim i As Integer

    CountVulns = 0
    
    'myError "UpdateVulns"
    Call UpdateVulns
    
    For i = DMG_SLASHING To DMG_LIGHTNING
        If IsVulnedTo(i) Then
            CountVulns = CountVulns + 1
        End If
    Next i
    
End Function
