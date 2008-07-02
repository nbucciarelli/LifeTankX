Attribute VB_Name = "Constants"
Option Explicit

Public Const FILTER_FULL_NAME = "DarkSide Filter"
Public Const FILTER_SHORT_NAME = "DSFilter"

Public Const PORTAL_TYPE_ENTER = &H4410&
Public Const PORTAL_TYPE_EXIT = &H408&

Public Const PS_BOLT = "8374"       ' Bolt Spell
Public Const PS_ARC = "8774"        ' Arc Spell
Public Const PS_STREAK = "374"      ' Streak spell

Public Const FOLDER_DATA = "Data"
Public Const FOLDER_SPELLS = "Spells"
Public Const FILE_MATERIALS = "Materials.dat"
Public Const FILE_DUNGEONS = "Dungeons.xml"
Public Const FILE_SPELLNAMES = "SpellNames.dat"

Public Const CSIDL_MYDOCUMENTS = &HC             'My Documents
Public Const CSIDL_APPDATA = &H1A           'Users App Data?
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_PERSONAL = &H5

Public Enum eNetMessages
    RETIRED_MSG_ATTACK = &H5E&          'a player attacks a monster
    
    MSG_GAME_EVENT = &HF7B0&            'Inbound Event
    MSG_GAME_ACTION = &HF7B1&           'Outbound Event
    MSG_ANIMATION = &HF74C&
    MSG_TOGGLE_VISIBILITY = &HF74B&     'Toggles portal animation or object visilibity (for war bolts)
    
    MSG_SET_CHAR_DWORD = &H2CD&         'Set Character DWORD
    MSG_SET_CHAR_QWORD = &H2CF&         'Set Character QWORD
    MSG_SET_CHAR_POSITION = &H2DB&      'Set Character Position
    
    RETIRED_MSG_SET_WIELDER_CONTAINER = &H22D&
    RETIRED_MSG_UPDATE_LAST_ATTACKER = &H23B&
    
    MSG_MOVE_INVENTORY_OBJECT = &HF74A& ' Move object into inventory.
    MSG_WIELD_OBJECT = &HF749&          'Wield Object
    MSG_ADJUST_STACK_SIZE = &H197&
    MSG_PLAYER_KILL = &H19E&
    MSG_SET_POSITION = &HF748&      'object position changed
    MSG_CREATE_OBJECT = &HF745&     'object created
    MSG_UPDATE_OBJECT = &HF7DB&     'object updated (after a tinker or salvage)
    MSG_REMOVE_OBJECT = &HF747&     'object removed from scene
    MSG_DESTROY_OBJECT = &H24&      'object destroyed
    
    RETIRED_MSG_VITAL_STAT_UPDATE = &H244&
    RETIRED_MSG_UPDATE_SKILL_XP = &H23E&
    RETIRED_MSG_UPDATE_ATTRIBS_XP = &H241&
    RETIRED_MSG_UPDATE_VITALS_XP = &H243&
    RETIRED_MSG_UPDATE_STATISTICS = &H237&
    
    MSG_APPLY_VISUAL_EFFECTS = &HF755&
    MSG_APPLY_SOUND_EFFECT = &HF750&
    MSG_JUMPING = &HF74E&
    MSG_PORTAL_SPACE = &HF751&
    
    RETIRED_MSG_SET_COVERAGE = &H229&
    
    MSG_CHAR_LIST = &HF658&
    MSG_LOGIN_CHAR = &HF746&
    
    MSG_MSG = &H2BB&                    'Creature Message
    MSG_MSG_RANGED = &H2BC&             'Creature Message (Ranged)
    
    MSG_TURBINE_CHAT = &HF7DE&          'Turbine Chat
    MSG_SERVER_CHAT = &HF7E0&           'Server Message
    
    MSG_SERVER_NAME = &HF7E1&           'Server Name
    
    RETIRED_MSG_LOCAL_CHAT = &H37&
    RETIRED_MSG_MOTD = &HF65A&
    
    RETIRED_MSG_START_3DMODE = &HF7C7&
    MSG_END_3DMODE = &HF653&
End Enum

Public Enum eGameActions
    ACT_OPTION = &H5&                   'Set Single Character Option
    ACT_USE_ITEM = &H36&                'Use Item
    ACT_CAST_SPELL = &H48&              'Cast Spell
    ACT_CAST_SPELL_ON = &H4A&           'Cast Spell on Object
    ACT_MATERIALIZE = &HA1&             'Materialize
                                        '0x01A1  Set Character Options
End Enum

Public Enum eGameEvents
    EV_LOGIN = &H13&
    EV_ALLEGIANCE_INFO = &H20&
    EV_INSERT_INVENTORY_ITEM = &H22&    'Insert Inventory Item
    EV_WEAR_ITEM = &H23&                'Wear Item
    EV_APPROACH_VENDOR = &H62&          'Opened a Vendor
    EV_FELLOWSHIP_QUIT = &HA3&
    EV_FELLOWSHIP_DISMISS = &HA4&
    EV_IDENTIFY_OBJECT = &HC9&
    EV_ID = &HC9&
    EV_GROUP_CHAT = &H147&
    EV_SET_PACK_CONTENT = &H196&        'sends number of items in pack
    EV_DROP_ITEM = &H19A&               'drop item from inventory
    EV_MELEE_ATTACK_COMPLETE = &H1A7&
    EV_MY_DEATH = &H1AC&
    EV_DEATH_MESSAGE = &H1AD&
    EV_INFLICT_MELEE_DMG = &H1B1&
    EV_RECEIVE_MELEE_DMG = &H1B2&
    EV_TARGET_EVADES_ATTACK = &H1B3&
    EV_PLAYER_EVADES_ATTACK = &H1B4&
    EV_UPDATE_TARGET_HEALTH = &H1C0&
    EV_READY = &H1C7&
    EV_ENTER_TRADE = &H1FD&
    EV_UPDATE_ITEM_MANA = &H264&
    EV_ACTION_FAILURE = &H28A&
    EV_DIRECT_CHAT = &H2BD&  'receiving a @tell
    
    EV_FELLOWSHIP_CREATE = &H2BE&
    EV_FELLOWSHIP_DISBANDS = &H2BF&
    EV_FELLOWSHIP_RECRUIT = &H2C0&
     RETIRED_EV_FELLOWSHIP_INVITATION = &H274&
    
     RETIRED_EV_ADD_SPELL = &H4C&
    EV_ADD_ENCHANTMENT = &H2C2&
    EV_REMOVE_ENCHANTMENT = &H2C3&
    EV_REMOVE_MULT_ENCHANTMENT = &H2C5&
    EV_REMOVE_ALL_ENCHANTMENT_SILENT = &H2C6&
    EV_REMOVE_ENCHANTMENT_SILENT = &H2C7&
    EV_REMOVE_MULT_ENCHANTMENT_SILENT = &H2C8&
End Enum

'0x00000800 @f: Fellowship broadcast
'0x00001000 @v: Patron to vassal
'0x00002000 @p: Vassal to patron
'0x00004000 @m: Follower to monarch
'0x01000000 @c: Covassal broadcast
'0x02000000 @a: Allegiance broadcast by monarch or speaker
Public Enum eGroupMask
    MSK_FELLOW = &H800&
    MSK_PATRON_TO_VASSAL = &H1000&
    MSK_VASSAL_TO_PATRON = &H2000&
    MSK_FOLLOW_TO_MONARCH = &H4000&
    MSK_COVASSAL = &H1000000
    MSK_ALLEGIANCE = &H2000000
End Enum
    
Public Enum eVitalAttributesId
    STAT_HEALTH = &H2&
    STAT_STAM = &H4&
    STAT_MANA = &H6&
End Enum
    
Public Enum eVitalStats
    VITAL_STAT_HEALTH = &H1&
    VITAL_STAT_STAM = &H3&
    VITAL_STAT_MANA = &H5&
End Enum

'Public Enum eVisualEffects
'    EF_WAR_LAUNCH = &H4&
'    EF_WAR_LAND = &H5&
'    EF_FESTER = &H25&
'    EF_FIRE_VULN = &H2B&
'    EF_PIERCE_VULN = &H2D&
'    EF_BLADE_VULN = &H2F&
'    EF_ACID_VULN = &H31&
'    EF_COLD_VULN = &H33&
'    EF_LIGHTNING_VULN = &H35&
'    EF_IMPERIL_BLUDG = &H37&
'    EF_YIELD = &H17&
'    EF_FIZZLE = &H50&
'    EF_EQUIP_ITEM = &H77&
'    EF_UNEQUIP_ITEM = &H78&
'End Enum

Public Enum eVisualEffects
    EF_WAR_LAUNCH = &H4&
    EF_WAR_LAND = &H5&
    EF_YIELD = &H17&
    EF_FESTER = &H26&
    EF_FIRE_VULN = &H2C&
    EF_PIERCE_VULN = &H2E&
    EF_BLADE_VULN = &H30&
    EF_ACID_VULN = &H32&
    EF_COLD_VULN = &H34&
    EF_LIGHTNING_VULN = &H36&
    EF_IMPERIL_BLUDG = &H38&
    EF_FIZZLE = &H51&
    EF_EQUIP_ITEM = &H78&
    EF_UNEQUIP_ITEM = &H79&
End Enum


'====================================================================
' Coverages
'====================================================================

Enum eItemCoverage
    COV_NONE = 0
    COV_HEAD = 1
    COV_UNDERWEAR_CHEST = 2
    COV_UNDERWEAR_GIRTH = 4
    COV_UNDERWEAR_UPPER_ARMS = 8
    COV_UNDERWEAR_LOWER_ARMS = &H10
    COV_HANDS = &H20
    COV_UNDERWEAR_UPPER_LEGS = &H40
    COV_UNDERWEAR_LOWER_LEGG = &H80
    COV_FEET = &H100
    COV_CHEST = &H200
    COV_GIRTH = &H400
    COV_UPPER_ARMS = &H800
    COV_LOWER_ARMS = &H1000
    COV_UPPER_LEGS = &H2000
    COV_LOWER_LEGS = &H4000
    
    COV_NECKLACE = &H8000&
    COV_BRACELET_RIGHT = &H10000
    COV_BRACELET_LEFT = &H20000
    COV_RING_RIGHT = &H40000
    COV_RING_LEFT = &H80000
    
    COV_WEAPON = &H100000
    COV_SHIELD = &H200000
    COV_BOW = &H400000
    COV_AMMO = &H800000
    COV_WAND = &H1000000
End Enum

Public Const MASK_JEWELRY = COV_NECKLACE Or COV_BRACELET_RIGHT Or COV_BRACELET_LEFT Or COV_RING_RIGHT Or COV_RING_LEFT
Public Const MASK_TOP = COV_CHEST Or COV_UPPER_ARMS Or COV_LOWER_ARMS
Public Const MASK_BOTTOM = COV_GIRTH Or COV_UPPER_LEGS Or COV_LOWER_LEGS
Public Const MASK_UNDERWEAR = COV_UNDERWEAR_CHEST Or COV_UNDERWEAR_GIRTH Or _
                                COV_UNDERWEAR_UPPER_ARMS Or COV_UNDERWEAR_LOWER_ARMS Or _
                                COV_UNDERWEAR_UPPER_LEGS Or COV_UNDERWEAR_LOWER_LEGG
Public Const MASK_UNDERWEAR_TOP = COV_UNDERWEAR_CHEST Or COV_UNDERWEAR_UPPER_ARMS Or COV_UNDERWEAR_LOWER_ARMS
Public Const MASK_UNDERWAER_BOTTOM = COV_UNDERWEAR_GIRTH Or COV_UNDERWEAR_UPPER_LEGS Or COV_UNDERWEAR_LOWER_LEGG
Public Const MASK_WIELDABLES = COV_WEAPON Or COV_BOW Or COV_WAND Or COV_AMMO Or COV_SHIELD

'0x3C melee UA weapon with no shield in attack stance
'0x3D Standing
'0x3E melee weapon with no shield in attack stance
'0x3F bow in attack stance
'0x40 melee weapon with shield in attack stance
'0x49 SpellCasting

Enum eAnimStances
    ANIM_UA_NO_SHIELD_ATK = &H3C&
    ANIM_STANDING = &H3D&
    ANIM_MELEE_NO_SHIELD_ATK = &H3E&
    ANIM_BOW_ATK = &H3F&
    ANIM_MELEE_SHIELD_ATK = &H40&
    ANIM_SPELLCASTING = &H49&
End Enum

Enum eAnimTypes
    ANIM_TYPE_GENERAL = 0
    ANIM_TYPE_MOVE_TO_OBJECT = &H6&
    ANIM_TYPE_MOVE_TO_POSITION = &H7&
    ANIM_TYPE_TURN_TO_OBJECT = &H8&
    ANIM_TYPE_TURN_TO_POSITION = &H9&
End Enum




