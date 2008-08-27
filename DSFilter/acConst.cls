VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Public Enum eVulnFlags
    FL_SLASHING = 1
    FL_BLUDGEONING = 2
    FL_PIERCING = 4
    FL_FIRE = 8
    FL_COLD = 16
    FL_ACID = 32
    FL_LIGHTNING = 64
    FL_IMPERIL = 128
End Enum

Public Enum eItemTypes
    ITEM_UNKNOWN = 0
    ITEM_MELEE_WEAPON
    ITEM_ARMOR
    ITEM_CLOTHING
    ITEM_JEWELRY
    ITEM_PYREAL
    ITEM_MISC
    ITEM_MISSILE_WEAPON
    ITEM_CONTAINER
    ITEM_FLETCHING
    ITEM_GEM
    ITEM_COMPS
    ITEM_SCROLL
    ITEM_BOOK
    ITEM_WAND
    ITEM_MANA_STONES
    ITEM_BASE_FLETCHING
    ITEM_BASE_COOKING
    ITEM_BASE_ALCHEMY
    ITEM_CRAFTED_FLETCHING
    ITEM_CRAFTED_COOKING
    ITEM_CRAFTED_ALCHEMY
    ITEM_SALVAGE
    ITEM_CORPSE
    ITEM_HEALING_KIT
    ITEM_ARROW
    ITEM_DOOR
    ITEM_FOOD
    ITEM_BUNDLE
    ITEM_KEY
    ITEM_PORTAL
    ITEM_TRADENOTE
    ITEM_PLANT
    ITEM_UST
    ITEM_LIFESTONE
    ITEM_LOCKPICK
    ITEM_FOCI
End Enum

Public Enum eUseTypes
    USETYPE_CONTAINER = &H1&
    USETYPE_INSCRIBABLE = &H2&
    USETYPE_CANNOT_PICKUP = &H4&
    USETYPE_PLAYER = &H8&
    USETYPE_CANNOT_SELECT = &H80&
    USETYPE_MERCHANT = &H200&
    USETYPE_DOOR = &H1000&
    USETYPE_LIFESTONE = &H4000&
    USETYPE_CORPSE = &H2000&
    USETYPE_HEALING_KIT = &H10000
    USETYPE_PORTAL = &H40000
    USETYPE_LOCKPICK = &H20000
    USETYPE_FOCI = &H800000
End Enum

Public Enum eCombatStates
    COMBATSTATE_NONE = -1
    COMBATSTATE_PEACE = 1
    COMBATSTATE_MELEE = 2
    COMBATSTATE_ARCHER = 4      'April fix
    COMBATSTATE_MAGIC = 8       'April fix
End Enum

Public Enum eMeleeSkills
    SKILL_NONE
    SKILL_AXE
    SKILL_DAGGER
    SKILL_MACE
    SKILL_SPEAR
    SKILL_STAFF
    SKILL_SWORD
    SKILL_UNARMED_COMBAT
    SKILL_NUM_MELEE_SKILLS
End Enum

Public Enum eArcherSkills
    SKILL_NONE
    SKILL_BOW
    SKILL_CROSSBOW
    SKILL_THROWN_WEAPONS
End Enum

Public Enum eEquipParts
    EQ_NONE = 0
    EQ_HEAD
    EQ_HANDS
    EQ_FEET
    EQ_TOP
    EQ_BOTTOM
    EQ_UNDERWEAR_TOP
    EQ_UNDERWEAR_BOTTOM
    EQ_WEAPON
    EQ_WAND
    EQ_SHIELD
End Enum

Enum eLevelId
    LEVELID_LEVEL = 1
End Enum

Enum eSkillId
    WEAPON_ANY_MELEE = -1
    WEAPON_ANY_MISSILE = -2
    
    SKILLID_AXE = 1
    SKILLID_BOW = 2
    SKILLID_CROSSBOW = 3
    SKILLID_DAGGER = 4
    SKILLID_MACE = 5
    SKILLID_MELEE_DEF = 6
    SKILLID_MISSILE_DEF = 7
    SKILLID_SPEAR = 9
    SKILLID_STAFF = &HA
    SKILLID_SWORD = &HB
    SKILLID_THROWN_WEAPONS = &HC
    SKILLID_UA = &HD
    SKILLID_ARCANE_LORE = &HE
    SKILLID_MAGIC_DEF = &HF
    SKILLID_MANA_CONV = &H10
    SKILLID_TINK_ITEM = &H12
    SKILLID_ASSESS_PERSON = &H13
    SKILLID_DECEPTION = &H14
    SKILLID_HEALING = &H15
    SKILLID_JUMP = &H16
    SKILLID_LOCKPICK = &H17
    SKILLID_RUN = &H18
    SKILLID_ASSESS_CREATURE = &H1B
    SKILLID_TINK_WEAPON = &H1C
    SKILLID_TINK_ARMOR = &H1D
    SKILLID_TINK_MAGIC = &H1E
    SKILLID_CREATURE = &H1F
    SKILLID_ITEM = &H20
    SKILLID_LIFE = &H21
    SKILLID_WAR = &H22
    SKILLID_LEADERSHIP = &H23
    SKILLID_FLETCHING = &H24
    SKILLID_ALCHEMY = &H26
    SKILLID_COOKING = &H27
End Enum


Public Enum eArmorTypes
    ARMORTYPE_UNKNOWN = 0
    ARMORTYPE_AMULI
    ARMORTYPE_CELDON
    ARMORTYPE_PLATEMAIL
    ARMORTYPE_YOROI
    ARMORTYPE_KOUJIA
    ARMORTYPE_LEATHER
    ARMORTYPE_SCALEMAIL
    ARMORTYPE_CHAINMAIL
    ARMORTYPE_COVENANT
    ARMORTYPE_CHIRAN
    ARMORTYPE_LORICA
    ARMORTYPE_NARIYID
    ARMORTYPE_DIFORSA
    ARMORTYPE_ALDURESSA
    ARMORTYPE_TENASSA
    ARMORTYPE_BUCKLER
    ARMORTYPE_KITESHIELD
    ARMORTYPE_ROUNDSHIELD
    ARMORTYPE_TOWERSHIELD
    ARMORTYPE_COVENANTSHIELD
    
    ARMORTYPE_ANY
End Enum

Public Enum eArmorCoverage
    ARMORCOVER_NONE = 0
    ARMORCOVER_HEAD = COV_HEAD
    ARMORCOVER_FEET = COV_FEET
    ARMORCOVER_HANDS = COV_HANDS
    ARMORCOVER_CHEST = COV_CHEST
    ARMORCOVER_GIRTH = COV_GIRTH
    
    ARMORCOVER_UPPER_ARMS = COV_UPPER_ARMS
    ARMORCOVER_LOWER_ARMS = COV_LOWER_ARMS
    ARMORCOVER_UPPER_LEGS = COV_UPPER_LEGS
    ARMORCOVER_LOWER_LEGS = COV_LOWER_LEGS
    
    ARMORCOVER_LEGGINGS = COV_UPPER_LEGS Or COV_LOWER_LEGS

    ARMORCOVER_TOP = MASK_TOP
    ARMORCOVER_BOTTOM = MASK_BOTTOM
    ARMORCOVER_SHIELD = COV_SHIELD
    
    ARMORCOVER_ANY = -1
End Enum

Public Enum eWeaponDamageFlags
    FL_DMG_SLASHING = 1
    FL_DMG_PIERCING = 2
    FL_DMG_BLUDGEON = 4
    FL_DMG_COLD = 8
    FL_DMG_FIRE = &H10
    FL_DMG_ACID = &H20
    FL_DMG_LIGHTNING = &H40
    
    FL_DMG_ANY = &H80
    FL_DMG_SLASHPIERCE = FL_DMG_PIERCING Or FL_DMG_SLASHING
End Enum

Public Enum eActionFailureErrors
    FAIL_TOO_BUSY = &H1D& '0x001D You 're too busy!
    FAIL_CHARGED_TOO_FAR = &H1C& '0x001C Charged Too Far
    FAIL_UNABLE_TO_MOVE_TO_OBJECT = &H39& '0x0039 Unable to move to object!
    FAIL_TOO_FATIGUED_TO_ATTACK = &H3F7& '0x03F7 You are too fatigued to attack!
    FAIL_OUT_OF_AMMUNITION = &H3F8& '0x03F8 You are out of ammunition!
    FAIL_MISSILE_ATTACK_MISFIRED = &H3F9& '0x03F9 Your missile attack misfired!
    FAIL_IMPOSSIBLE_SPELL_PATH = &H3FA& '0x03FA You've attempted an impossible spell path!
    FAIL_DONT_KNOW_SPELL = &H3FE& '0X03FE You don't know that spell!
    FAIL_INCORRECT_TARGET_TYPE = &H3FF& '0X03FF Incorrect target type
    FAIL_NOT_ENOUGH_COMPONENTS = &H400& '0x0400 You don't have all the components for this spell.
    FAIL_NOT_ENOUGH_MANA = &H401& '0x0401 You don't have enough Mana to cast this spell.
    FAIL_SPELL_FIZZLE = &H402& '0x0402 Your spell fizzled.
    FAIL_NO_SPELL_TARGET = &H403& '0x0403 Your spell's target is missing!
    FAIL_PROJECTILE_MISLAUNCHED = &H404& '0x0404 Your projectile spell mislaunched!
    FAIL_SOLVED_QUEST_TOO_RECENTLY = &H43E& '0x043E You have solved this quest too recently!

    'not from Decal Protocol
    FAIL_INVALID_COMMAND = 38       'This is not a valid console command
    FAIL_UNPREPARED_TO_CAST = 1034  'Your are unprepared to cast that spell
    FAIL_NOT_ENOUGH_MANA_ON_ITEM = &H4C8&
    
    
    'Custom Error IDs
    FAIL_NOT_READY = -1
    FAIL_CAN_ONLY_MOVE_ONE = -2
    FAIL_CANT_OPEN_BODY = -3
    FAIL_ACTION_CANCELED = -4
    FAIL_ALREADY_IN_USE = -6
    FAIL_CANNOT_BE_USED = -7
    FAIL_BUSY = -8
End Enum

Public Enum eGameServer
    SV_WINTERSEBB = 0
    SV_MORNINGTHAW
    SV_HARVESTGAIN
    SV_FROSTFELL
    SV_SOLCLAIM
    SV_THISTLEDOWN
    SV_LEAFCULL
    SV_DARKTIDE
    SV_VERDANTINE
    NUM_AC_SERVERS
    SV_ANY
    SV_NONE
End Enum

'############################################################

Private Const TAG_MATERIAL_ID = "id"
Private Const TAG_MATERIAL_NAME = "name"

Private Const TAG_SPELL_NAME = "spl"
Private Const TAG_SPELL_ID = "id"

'############################################################

Private m_colSkills As New colDatas
Private m_colWeaponSkills As New colDatas
Private m_colArmorTypes As New colDatas
Private m_colArmorCoverage As New colDatas
Private m_colMaterials As New colDatas
Private m_colWeaponDamages As New colDatas
Private m_colItemTypes As New colDatas
Private m_colLevel As New colDatas

Private m_colDungeons As New Collection
Private m_dicSpellNames As New Dictionary

Private m_db As New DataFile

Private Sub Class_Initialize()
    Call LoadSkills
    Call LoadLevel
    Call LoadWeaponsSkills
    Call LoadWeaponDamages
    Call LoadArmors
    Call LoadArmorCoverage
    Call LoadItemTypes
    Call LoadDungeons
    Call LoadSpellNames(GetDataFolder & "\" & FOLDER_DATA & "\" & FOLDER_SPELLS & "\" & FILE_SPELLNAMES)
End Sub

Private Sub Class_Terminate()
    Set m_colSkills = Nothing
    Set m_colLevel = Nothing
    Set m_colWeaponSkills = Nothing
    Set m_colArmorCoverage = Nothing
    Set m_colArmorTypes = Nothing
    Set m_colMaterials = Nothing
    Set m_colWeaponDamages = Nothing
    Set m_colItemTypes = Nothing
    Set m_dicSpellNames = Nothing
    Set m_db = Nothing
    Set m_colDungeons = Nothing
End Sub

Public Property Get Materials() As colDatas
    Set Materials = m_colMaterials
End Property

Public Property Get Skills() As colDatas
    Set Skills = m_colSkills
End Property
Public Property Get Level() As colDatas
    Set Level = m_colLevel
End Property

Public Property Get WeaponSkills() As colDatas
    Set WeaponSkills = m_colWeaponSkills
End Property

Public Property Get Armors() As colDatas
    Set Armors = m_colArmorTypes
End Property

Public Property Get ArmorCoverage() As colDatas
    Set ArmorCoverage = m_colArmorCoverage
End Property

Public Property Get ItemTypes() As colDatas
    Set ItemTypes = m_colItemTypes
End Property

Public Property Get WeaponDamages() As colDatas
    Set WeaponDamages = m_colWeaponDamages
End Property

Private Sub LoadSkills()
    With m_colSkills
        .Clear
        .Add SKILLID_AXE, "Axe"
        .Add SKILLID_BOW, "Bow"
        .Add SKILLID_CROSSBOW, "Crossbow"
        .Add SKILLID_DAGGER, "Dagger"
        .Add SKILLID_MACE, "Mace"
        .Add SKILLID_MELEE_DEF, "Melee Defense"
        .Add SKILLID_MISSILE_DEF, "Missile Def"
        .Add SKILLID_SPEAR, "Spear"
        .Add SKILLID_STAFF, "Staff"
        .Add SKILLID_SWORD, "Sword"
        .Add SKILLID_THROWN_WEAPONS, "Thrown Weapons"
        .Add SKILLID_UA, "Unarmed Combat"
        .Add SKILLID_ARCANE_LORE, "Arcane Lore"
        .Add SKILLID_MAGIC_DEF, "Magic Defense"
        .Add SKILLID_MANA_CONV, "Mana Conversion"
        .Add SKILLID_TINK_ITEM, "Item Tinkering"
        .Add SKILLID_ASSESS_PERSON, "Assess Person"
        .Add SKILLID_DECEPTION, "Deception"
        .Add SKILLID_HEALING, "Healing"
        .Add SKILLID_JUMP, "Jump"
        .Add SKILLID_LOCKPICK, "Lockpick"
        .Add SKILLID_RUN, "Run"
        .Add SKILLID_ASSESS_CREATURE, "Assess Creature"
        .Add SKILLID_TINK_WEAPON, "Weapon Tinkering"
        .Add SKILLID_TINK_ARMOR, "Armor Tinkering"
        .Add SKILLID_TINK_MAGIC, "Magic Tinkering"
        .Add SKILLID_CREATURE, "Creature Enchantment"
        .Add SKILLID_ITEM, "Item Enchantment"
        .Add SKILLID_LIFE, "Life Magic"
        .Add SKILLID_WAR, "War Magic"
        .Add SKILLID_LEADERSHIP, "Leadership"
        .Add SKILLID_ALCHEMY, "Alchemy"
        .Add SKILLID_FLETCHING, "Fletching"
        .Add SKILLID_COOKING, "Cooking"
    End With
End Sub

Private Sub LoadLevel()
    With m_colLevel
        .Clear
        .Add LEVELID_LEVEL, "Level"
    End With
End Sub

Private Sub LoadWeaponsSkills()
    With m_colWeaponSkills
        .Clear
        .Add WEAPON_ANY_MELEE, "Any", ITEM_MELEE_WEAPON
        .Add WEAPON_ANY_MISSILE, "Any", ITEM_MISSILE_WEAPON
        .Add SKILLID_AXE, "Axe", ITEM_MELEE_WEAPON
        .Add SKILLID_BOW, "Bow", ITEM_MISSILE_WEAPON
        .Add SKILLID_CROSSBOW, "Xbow", ITEM_MISSILE_WEAPON
        .Add SKILLID_DAGGER, "Dagger", ITEM_MELEE_WEAPON
        .Add SKILLID_MACE, "Mace", ITEM_MELEE_WEAPON
        .Add SKILLID_SPEAR, "Spear", ITEM_MELEE_WEAPON
        .Add SKILLID_STAFF, "Staff", ITEM_MELEE_WEAPON
        .Add SKILLID_SWORD, "Sword", ITEM_MELEE_WEAPON
        .Add SKILLID_THROWN_WEAPONS, "TW", ITEM_MISSILE_WEAPON
        .Add SKILLID_UA, "UA", ITEM_MELEE_WEAPON
    End With
End Sub
Private Sub LoadArmors()
    With m_colArmorTypes
        .Clear
        .Add ARMORTYPE_ANY, "Any"
        .Add ARMORTYPE_AMULI, "Amuli"
        .Add ARMORTYPE_CELDON, "Celdon"
        .Add ARMORTYPE_PLATEMAIL, "Platemail"
        .Add ARMORTYPE_YOROI, "Yoroi"
        .Add ARMORTYPE_KOUJIA, "Koujia"
        .Add ARMORTYPE_LEATHER, "Leather"
        .Add ARMORTYPE_SCALEMAIL, "Scalemail"
        .Add ARMORTYPE_CHAINMAIL, "Chainmail"
        .Add ARMORTYPE_COVENANTSHIELD, "Covenant Shield"
        .Add ARMORTYPE_COVENANT, "Covenant"
        .Add ARMORTYPE_CHIRAN, "Chiran"
        .Add ARMORTYPE_LORICA, "Lorica"
        .Add ARMORTYPE_NARIYID, "Nariyid"
        .Add ARMORTYPE_DIFORSA, "Diforsa"
        .Add ARMORTYPE_ALDURESSA, "Alduressa"
        .Add ARMORTYPE_TENASSA, "Tenassa"
        .Add ARMORTYPE_BUCKLER, "Buckler"
        .Add ARMORTYPE_KITESHIELD, "Kite Shield"
        .Add ARMORTYPE_ROUNDSHIELD, "Round Shield"
        .Add ARMORTYPE_TOWERSHIELD, "Tower Shield"
    End With
End Sub

Private Sub LoadArmorCoverage()
    With m_colArmorCoverage
        .Clear
        .Add ARMORCOVER_ANY, "Any"
        .Add ARMORCOVER_SHIELD, "Shield"
        .Add ARMORCOVER_HEAD, "Head"
        .Add ARMORCOVER_FEET, "Feet"
        .Add ARMORCOVER_HANDS, "Hands"
        .Add ARMORCOVER_CHEST, "Chest"
        .Add ARMORCOVER_GIRTH, "Girth"
        .Add ARMORCOVER_LEGGINGS, "Leggings"
        .Add ARMORCOVER_TOP, "Top"
        .Add ARMORCOVER_BOTTOM, "Bottom"
        .Add ARMORCOVER_UPPER_ARMS, "Upper Arms"
        .Add ARMORCOVER_LOWER_ARMS, "Lower Arms"
        .Add ARMORCOVER_UPPER_LEGS, "Upper Legs"
        .Add ARMORCOVER_LOWER_LEGS, "Lower Legs"
    End With
End Sub

Private Sub LoadWeaponDamages()
    With m_colWeaponDamages
        .Clear
        .Add FL_DMG_ANY, "Any"
        .Add FL_DMG_SLASHING, "Slash"
        .Add FL_DMG_PIERCING, "Pierce"
        .Add FL_DMG_SLASHPIERCE, "S/P"
        .Add FL_DMG_BLUDGEON, "Bludgeon"
        .Add FL_DMG_COLD, "Frost"
        .Add FL_DMG_FIRE, "Fire"
        .Add FL_DMG_ACID, "Acid"
        .Add FL_DMG_LIGHTNING, "Electric"
    End With
End Sub

Private Sub LoadItemTypes()
    With m_colItemTypes
        .Clear
        .Add ITEM_UNKNOWN, "Unknown Item"
        .Add ITEM_MELEE_WEAPON, "Melee Weapon"
        .Add ITEM_ARMOR, "Armor"
        .Add ITEM_CLOTHING, "Clothing"
        .Add ITEM_JEWELRY, "Jewelry"
        .Add ITEM_PYREAL, "Pyreal"
        .Add ITEM_MISC, "Misc"
        .Add ITEM_MISSILE_WEAPON, "Missile Weapon"
        .Add ITEM_CONTAINER, "Container"
        .Add ITEM_FLETCHING, "Fletching"
        .Add ITEM_GEM, "Gem"
        .Add ITEM_COMPS, "Component"
        .Add ITEM_SCROLL, "Scroll"
        .Add ITEM_BOOK, "Book"
        .Add ITEM_WAND, "Magic Wand"
        .Add ITEM_MANA_STONES, "Mana Stone"
        .Add ITEM_BASE_FLETCHING, "Base Fletching"
        .Add ITEM_BASE_COOKING, "Base Cooking"
        .Add ITEM_BASE_ALCHEMY, "Base Alchemy"
        .Add ITEM_CRAFTED_FLETCHING, "Crafted Fletching"
        .Add ITEM_CRAFTED_COOKING, "Crafter Cooking"
        .Add ITEM_CRAFTED_ALCHEMY, "Crafter Alchemy"
        .Add ITEM_SALVAGE, "Salvage"
        .Add ITEM_CORPSE, "Corpse"
        .Add ITEM_HEALING_KIT, "Healing Kit"
        .Add ITEM_ARROW, "Arrow"
        .Add ITEM_DOOR, "Door"
        .Add ITEM_FOOD, "Food"
        .Add ITEM_BUNDLE, "Bundle"
        .Add ITEM_KEY, "Key"
        .Add ITEM_PORTAL, "Portal"
        .Add ITEM_TRADENOTE, "Tradenote"
        .Add ITEM_PLANT, "Plant"
        .Add ITEM_UST, "Ust"
        .Add ITEM_LIFESTONE, "Lifestone"
        .Add ITEM_LOCKPICK, "Lockpick"
        .Add ITEM_FOCI, "Foci"
    End With
End Sub

Public Function GetSkillName(ByVal lId As Long) As String
    GetSkillName = m_colSkills(lId)
End Function

Public Function GetLevelName(ByVal lId As Long) As String
    GetLevelName = m_colLevel(lId)
End Function

Public Function GetWeaponSkillName(ByVal lId As Long) As String
    GetWeaponSkillName = m_colWeaponSkills(lId)
End Function

Public Function GetItemTypeName(ByVal lItemType As Long) As String
    GetItemTypeName = m_colItemTypes(lItemType)
End Function

Friend Sub Save()
    Call SaveMaterialsList
End Sub

Friend Function LoadMaterialsList(ByVal sPath As String) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim dat As clsDataEntry

    If m_db.Load(sPath) Then
        For Each dat In m_db
            If dat.ParamExist(TAG_MATERIAL_ID) And dat.ParamExist(TAG_MATERIAL_NAME) Then
                Call m_colMaterials.Add(CLng(Val(dat.Param(TAG_MATERIAL_ID))), dat.Param(TAG_MATERIAL_NAME))
            End If
        Next dat
        myDebug "Materials List loaded successfully from " & sPath
        bRet = True
    Else
        myError "acConst.LoadMaterialsList - Failed to load " & sPath
        bRet = False
    End If
    
Fin:
    LoadMaterialsList = bRet
    Exit Function
ErrorHandler:
    myError "acConst.LoadMaterialsList(" & sPath & ") - " & Err.Description
    bRet = False
    Resume Fin
End Function

Friend Function SaveMaterialsList(Optional ByVal sPath As String = "") As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim dat As clsDataEntry
    Dim oData As DataItem

    Call m_db.ResetData
    
    If sPath = "" Then
        sPath = m_db.FileName
    End If
    
    For Each oData In m_colMaterials
        Set dat = New clsDataEntry
        Call dat.AddParam(TAG_MATERIAL_ID, oData.Key)
        Call dat.AddParam(TAG_MATERIAL_NAME, oData.Val)
        Call m_db.AddData(dat)
    Next oData
    
    If m_db.Save(sPath) Then
        myDebug "acConst.SaveMaterialsList - List saved to " & sPath
        bRet = True
    Else
        myError "acConst.SaveMaterialsList - Failed to save list to " & sPath
        bRet = False
    End If
    
Fin:
    SaveMaterialsList = bRet
    Set oData = Nothing
    Set dat = Nothing
    Exit Function
ErrorHandler:
    myError "acConst.SaveMaterialsList(" & sPath & ") - " & Err.Description
    bRet = False
    Resume Fin
End Function

Friend Sub CheckNewMaterial(ByVal lMaterialId As Long, ByVal sMaterialName As String)
    If Not m_colMaterials.Exists(lMaterialId) Then
        myDebug "Added Material ID #" & lMaterialId & " from " & sMaterialName
        Call m_colMaterials.Add(lMaterialId, sMaterialName)
    End If
End Sub

Friend Function FindDungeon(ByVal lDungeonId As Long, Optional ByRef oDungeonOut As DungeonData) As Boolean
On Error GoTo ErrorHandler

    Set oDungeonOut = m_colDungeons(CStr(lDungeonId))
    FindDungeon = True

Fin:
    Exit Function
ErrorHandler:
    FindDungeon = False
    Resume Fin
End Function

Friend Function LoadDungeons() As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As New DOMDocument
    
    bRet = False
    Set m_colDungeons = New Collection
    
    myDebug "Loading Dungeons: " & GetDataFolder & "\" & FOLDER_DATA & "\" & FILE_DUNGEONS
    
    If xmlDoc.Load(GetDataFolder & "\" & FOLDER_DATA & "\" & FILE_DUNGEONS) Then
        For Each xmlNode In xmlDoc.documentElement.childNodes
            Dim nodeAttrib As IXMLDOMAttribute
            
            Dim oDungeon As DungeonData
            Set oDungeon = New DungeonData
            
            For Each nodeAttrib In xmlNode.Attributes
                Select Case LCase(nodeAttrib.nodeName)
                
                    Case "id"
                        oDungeon.Id = CLng("&H" + nodeAttrib.Value)
                        
                    Case "name"
                        oDungeon.Name = nodeAttrib.Value
                        
                    Case "coords"
                        oDungeon.Info = nodeAttrib.Value
                        
                End Select
            Next nodeAttrib
            
            If Not FindDungeon(oDungeon.Id) Then
                'myDebug "Added Dungeon: " & oDungeon.Name & " (" & Hex(oDungeon.Id) & ")"
                Call m_colDungeons.Add(oDungeon, CStr(oDungeon.Id))
            End If
            
            Set oDungeon = Nothing
        Next xmlNode
        bRet = True
    Else
        myError "acConst.LoadDungeons - Failed to load XML file"
    End If
    
Fin:
    LoadDungeons = bRet
    Exit Function
ErrorHandler:
    myError "acConst.LoadDungeons - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function GetServerName(ByVal ServerID As eGameServer) As String
Dim sRet As String
    
    Select Case ServerID
        Case SV_WINTERSEBB
            sRet = "Wintersebb"
            
        Case SV_MORNINGTHAW
            sRet = "Morningthaw"
    
        Case SV_HARVESTGAIN
            sRet = "Harvestgain"
            
        Case SV_FROSTFELL
            sRet = "Frostfell"
            
        Case SV_SOLCLAIM
            sRet = "Solclaim"
            
        Case SV_THISTLEDOWN
            sRet = "Thistledown"
    
        Case SV_LEAFCULL
            sRet = "Leafcull"
            
        Case SV_DARKTIDE
            sRet = "Darktide"
            
        Case SV_VERDANTINE
            sRet = "Verdantine"
            
        Case SV_ANY
            sRet = "Any"
        
        Case Else
            sRet = "Unknown Server " & ServerID
            
    End Select

    GetServerName = sRet
End Function

Public Function GetShortServerName(ByVal ServerID As eGameServer) As String
Dim sRet As String
    
    Select Case ServerID
        Case SV_WINTERSEBB
            sRet = "WE"
            
        Case SV_MORNINGTHAW
            sRet = "MT"
    
        Case SV_HARVESTGAIN
            sRet = "HG"
            
        Case SV_FROSTFELL
            sRet = "FF"
            
        Case SV_SOLCLAIM
            sRet = "SC"
            
        Case SV_THISTLEDOWN
            sRet = "TD"
    
        Case SV_LEAFCULL
            sRet = "LC"
            
        Case SV_DARKTIDE
            sRet = "DT"
            
        Case SV_VERDANTINE
            sRet = "VT"
            
        Case SV_ANY
            sRet = "ANY"
        
        Case Else
            sRet = "UNKN" & ServerID
            
    End Select

    GetShortServerName = sRet
End Function

Public Function GetVBServerName(ByVal ServerID As eGameServer) As String
Dim sRet As String
    
    Select Case ServerID
        Case SV_WINTERSEBB
            sRet = "SV_WINTERSEBB"
            
        Case SV_MORNINGTHAW
            sRet = "SV_MORNINGTHAW"
    
        Case SV_HARVESTGAIN
            sRet = "SV_HARVESTGAIN"
            
        Case SV_FROSTFELL
            sRet = "SV_FROSTFELL"
            
        Case SV_SOLCLAIM
            sRet = "SV_SOLCLAIM"
            
        Case SV_THISTLEDOWN
            sRet = "SV_THISTLEDOWN"
    
        Case SV_LEAFCULL
            sRet = "SV_LEAFCULL"
            
        Case SV_DARKTIDE
            sRet = "SV_DARKTIDE"
            
        Case SV_VERDANTINE
            sRet = "SV_VERDANTINE"
            
        Case SV_ANY
            sRet = "SV_ANY"
        
        Case SV_NONE
            sRet = "SV_NONE"
        
        Case Else
            sRet = "Unknown Serverid " & ServerID
            
    End Select

    GetVBServerName = sRet
End Function

Public Function GetServerIdByName(ByVal ServerName As String) As eGameServer
    Dim iRet As eGameServer
    
    ServerName = LCase(ServerName)
    
    Select Case ServerName
        Case "wintersebb"
            iRet = SV_WINTERSEBB
            
        Case "morningthaw"
            iRet = SV_MORNINGTHAW
    
        Case "harvestgain"
            iRet = SV_HARVESTGAIN
            
        Case "frostfell"
            iRet = SV_FROSTFELL
            
        Case "solclaim"
            iRet = SV_SOLCLAIM
            
        Case "thistledown"
            iRet = SV_THISTLEDOWN
    
        Case "leafcull"
            iRet = SV_LEAFCULL
            
        Case "darktide"
            iRet = SV_DARKTIDE
            
        Case "verdantine"
            iRet = SV_VERDANTINE
        
        Case Else
            iRet = SV_NONE
            
    End Select
    
    'Assume it's darktide if server name couldn't be determined...
    If iRet = SV_NONE Then 'And InStr(1, ServerName, "darktide") Then
        iRet = SV_DARKTIDE
    End If
    
    GetServerIdByName = iRet
    
End Function

Public Function GetSpellName(ByVal lSpellID As Long) As String
On Error GoTo ErrorHandler

    Dim sRet As String
    If m_dicSpellNames.Exists(lSpellID) Then
        sRet = m_dicSpellNames(lSpellID)
    Else
        sRet = ""
    End If
    
Fin:
    GetSpellName = sRet
    Exit Function
ErrorHandler:
    sRet = ""
    myError "acConst.GetSpellName(" & lSpellID & ") - " & Err.Description
    Resume Fin
End Function

'Load the SpellNames.dat (data file with only spells ID/Name pairs
Private Function LoadSpellNames(ByVal sPath As String) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    Dim db As New DataFile
    Dim dat As clsDataEntry
    Dim i As Integer
       
    If db.Load(sPath) Then
        For Each dat In db
            i = i + 1
            If dat.ParamExist(TAG_SPELL_ID) And dat.ParamExist(TAG_SPELL_NAME) Then
                Dim lSpellID As Long
                lSpellID = Val(dat.Param(TAG_SPELL_ID))
                Call m_dicSpellNames.Add(lSpellID, dat.Param(TAG_SPELL_NAME))
            End If
        Next dat
        
        myDebug "LoadSpellNames : loaded " & m_dicSpellNames.Count & "/" & i & " spellnames"
        bRet = True
    Else
        myError "acConst.LoadSpellNames : failed to load SpellNames from " & sPath
        bRet = False
    End If
    
Fin:
    Set db = Nothing
    Set dat = Nothing
    LoadSpellNames = bRet
    Exit Function
ErrorHandler:
    bRet = False
    myError "acConst.LoadSpellNames(" & sPath & ") - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

