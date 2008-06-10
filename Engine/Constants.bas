Attribute VB_Name = "Constants"
'----------------------------------------------
'Shared Module
'----------------------------------------------

Option Explicit

'=============================================================
'                  Paths & Filenames
'=============================================================

Public Const PATH_DATA = "Data"
Public Const PATH_LOOT = "Loot"
Public Const PATH_LOGS = "Logs"
Public Const PATH_SPELLS = PATH_DATA & "\Spells"
Public Const PATH_SOUNDS = PATH_DATA & "\Sounds"
Public Const PATH_ROUTES = PATH_DATA & "\Routes"

Public Const FOLDER_PROFILE_LOGS = "Logs"
Public Const FOLDER_PROFILE_MACRO = "Macro"
    Public Const FILE_MACRO_CONFIG = "MacroCfg.ini"
Public Const FOLDER_PROFILE_BUFFS = "Buffs"
    Public Const FILE_BUFFS_CONFIG = "BuffsCfg.ini"
Public Const PATH_PROFILE_LOOT = PATH_DATA & "\Loot"
    Public Const FILE_LOOT_CONFIG = "LootCfg.ini"

Public Const FILE_CONFIG = "LifeTank.ini"
Public Const FILE_SHARED_CONFIG = "Shared.ini"
Public Const FILE_EXCEPTIONS = "Exceptions.xml"

Public Const FILE_AUTORESPONSE = "AutoRespond.xml"

Public Const FILE_EXT_ROUTE = "rte" ' *.rte

Public Const FILE_DEATHMESSAGES = "DeathMessages.txt"

Public Const FILE_MATERIALS = "Materials.dat"
Public Const FILE_MONSTERS = "Monsters.dat"
Public Const FILE_LOOT_FILTERS = "LootFilters.dat"

Public Const FILE_ITEM_BANES = "ItemBanes.dat"
Public Const FILE_BUFF_LIST = "Buffs.ini"

Public Const FILE_FELLOW_BAN_LIST = "BanList.dat"
Public Const FILE_FELLOW_FRIENDS_LIST = "FellowFriendsList.dat"

Public Const FILE_ITEMS_TO_PICKUP_LIST = "ItemsToPickup.dat"
Public Const FILE_ITEMS_TO_IGNORE_LIST = "ItemsToIgnore.dat"
Public Const FILE_CORPSES_TO_LOOT_LIST = "CorpsesToLoot.dat"

Public Const FILE_HEALITEMS_LIST = "Healitems.dat"
Public Const FILE_STAMITEMS_LIST = "Stamitems.dat"
Public Const FILE_EMERGITEMS_LIST = "Emergitems.dat"

Public Const FILE_ANTIBAN_FRIENDS_LIST = "FriendsList.dat"
Public Const FILE_ANTIBAN_ALLEGIANCES_LIST = "AllegianceFriends.dat"

Public Const FILE_DETECT_LIST = "DetectList.dat"

Public Const FILE_WARSPELLS = "WarSpells.dat"
Public Const FILE_ITEMSPELLS = "ItemSpells.dat"
Public Const FILE_OTHERBUFFS = "BuffsOther.dat"
Public Const FILE_SELFBUFFS = "BuffsSelf.dat"
Public Const FILE_DEBUFFSPELLS = "Debuffs.dat"
Public Const FILE_VITALSPELLS = "MacroSpells.dat"
Public Const FILE_SPELLNAMES = "SpellNames.dat"

Public Const CSIDL_MYDOCUMENTS = &HC             'My Documents
Public Const CSIDL_APPDATA = &H1A           'Users App Data?
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_PERSONAL = &H5

'=============================================================
'                       Sounds
'=============================================================

Public Const SOUND_ALARM = PATH_SOUNDS & "\alarm.wav"
Public Const SOUND_ENVOY = PATH_SOUNDS & "\envoyradar.wav"
Public Const SOUND_TELL = PATH_SOUNDS & "\tell.wav"
Public Const SOUND_OPENCHAT = PATH_SOUNDS & "\openchat.wav"
Public Const SOUND_DETECT = PATH_SOUNDS & "\detect.wav"
Public Const SOUND_BACKPACK_FULL = PATH_SOUNDS & "\backpackfull.wav"
Public Const SOUND_FELLOW_DISBAND = PATH_SOUNDS & "\fellowdisband.wav"
Public Const SOUND_FELLOW_DEAD = PATH_SOUNDS & "\fellowdead.wav"
Public Const SOUND_RING = PATH_SOUNDS & "\ringring.wav"
Public Const SOUND_EMERGENCY = PATH_SOUNDS & "\emerg.wav"
Public Const SOUND_DEATH = PATH_SOUNDS & "\death.wav"
Public Const SOUND_RARE = PATH_SOUNDS & "\rare.wav"

'=============================================================
'                  Database Param Tags
'=============================================================

Public Const TAG_ID = "id"
Public Const TAG_VALUE = "value"
Public Const ID_ROUTE_NAME = "Route"
Public Const ID_ROUTE_NAVTYPE = "RouteNavType"
Public Const TAG_MACRO_PROFILE = "MacroProfile"
Public Const TAG_BUFF_PROFILE = "BuffProfile"
Public Const TAG_LOOT_PROFILE = "LootProfile"
Public Const TAG_FILTER_TYPE = "type"
Public Const TAG_FILTER_ENABLED = "enabled"

'=============================================================
'                         Profiles
'=============================================================

Public Const PROFILE_DEFAULT = "Default"
Public Const PROFILES_FOLDER = PATH_DATA & "\Profiles"
Public Const DEFAULT_PROFILE_FOLDER = PROFILES_FOLDER & "\" & PROFILE_DEFAULT
Public Const PROFILE_BUFFS_FOLDER = "Buffs"


'=============================================================
'                       Strings
'=============================================================

Public Const MACRO_REMOTE_COMMAND_TAG = "#cmdmacro"
Public Const MACRO_IRC_TAG = "LTx_"
Public Const STR_ITEM_PLAT = "Platinum Scarab"
Public Const STR_ITEM_PYREAL_SCARAB = "Pyreal Scarab"
Public Const STR_ITEM_DIAMOND_SCARAB = "Diamond Scarab"
Public Const STR_ITEM_STAM_POTION = "Stamina Elixir"
Public Const STR_ITEM_HEALTH_POTION = "Health Elixir"
Public Const STR_ITEM_TAPER = "Prismatic Taper"
Public Const STR_ITEM_HEALING_KIT = "Healing Kit"
Public Const STR_ITEM_MANA_CHARGE = "*Mana Charge"
Public Const STR_ITEM_UST = "Ust"

'=============================================================
'                       Spell Family Names
'=============================================================

Public Const SPELL_BD = "Blood Drinker"
Public Const SPELL_HS = "Heart Seeker"
Public Const SPELL_SK = "Swift Killer"
Public Const SPELL_DEF = "Defender"
Public Const SPELL_HERMETIC_LINK = "Hermetic Link"
Public Const SPELL_SPIRIT_DRINKER = "Spirit Drinker"

Public Const SPELL_IMPEN = "Impenetrability"

Public Const SPELL_BANE_ACID = "Acid Bane"
Public Const SPELL_BANE_COLD = "Frost Bane"
Public Const SPELL_BANE_LIGHTNING = "Lightning Bane"
Public Const SPELL_BANE_FIRE = "Flame Bane"

Public Const SPELL_BANE_BLUDEONING = "Bludgeon Bane"
Public Const SPELL_BANE_SLASHING = "Blade Bane"
Public Const SPELL_BANE_PIERCING = "Piercing Bane"


'=============================================================
'                       Misc
'=============================================================

Public Const ANTI_LOGOUT_TIMER_INTERVAL = 300 'send key every 5 minutes
Public Const FELLOW_NOTIFY_INTERVAL = 10 'minutes (FIXME 15 mins)
Public Const M_PI = 3.14159265358979
Public Const BLACKLIST_TIME = 30    'stay blacklisted 30 seconds
Public Const BLACKLIST_MAXCOUNT = 3 '3 blacklist hit = perma blacklisted

' More Icons
Global Const MOVE_UP_ICON = &H60028FC
Global Const MOVE_DOWN_ICON = &H60028FD
Global Const DEL_ICON = &H6005E6A


'=============================================================
'                   Melee Restam
'=============================================================

Public Const RESTAM_MIN_STAM_PERCENT = 30
Public Const RESTAM_MAX_STAM_PERCENT = 90
Public Const RESTAM_MAX_STAM_PERCENT_MELEE = 98

'=============================================================
'                 Extra acObject Infos
'=============================================================
Public Const CORPSE_TIMER = 290     'just under 5 mins

Enum eExtraObjectInfo
    B_ENABLED = 0           'Is monster checked in the Monsters list?
    B_CAN_BE_IMPERILED      'Can this monster be imperiled?
    B_CAN_BE_VULNED         'Can this monster be vulned?
    B_CAN_BE_YIELDED        'Can this monster be yielded?
    B_DANGEROUS             'Is this target dangerous?
    B_WAR_CHECK             'Fire no War spells check
    INT_BLIST_TIME          'Time this target is blacklisted for
    INT_MISSCOUNT           'Number of times we misfired on this target
    INT_BLISTCOUNT          'Number of times this critter has been blacklisted
    INT_YIELD_TRYS_LEFT     'Number of trys left to yield this target
    B_LOOTED                'Has this corpse been looted by the macro already?
    B_HASRARE               'This corpse has a rare on it
    B_MACO_PICKUP           'Has this item been picked up by the macro?
    L_TIME                  'Time to expire this object?
    INT_DELETE              'Should we ask DS Filter to delete this object
    INT_SALVAGECOUNT        'Number of Times we have attempted to salvage this item

    NUM_EXTRA_OBJECT_INFO
End Enum

'=============================================================
'                   Rebuff Modes
'=============================================================
Public Enum eRebuffModes
    REBUFF_FULL = 0
    REBUFF_CONTINUOUS
End Enum

'=============================================================
'           Macro States/Actions/Types/Modes
'=============================================================

Public Enum eCombatStates
    COMBATSTATE_NONE = -1
    COMBATSTATE_PEACE = 1
    COMBATSTATE_MELEE = 2
    COMBATSTATE_ARCHER = 4      'April fix
    COMBATSTATE_MAGIC = 8       'April fix
End Enum

Enum eActionType
    ACT_NONE
    ACT_YIELDMONSTER
    ACT_WAR_PRIMARY
    ACT_ARC
    ACT_VULN
    ACT_IMPERIL
    ACT_REVITALIZE
    ACT_STAMTOMANA
    ACT_HEALTH_TO_MANA
    ACT_ATTACK
End Enum


Public Const ID_WAND = "wand"
Public Const ID_WEAPON = "weapon"
Public Const ID_SHIELD = "shield"
Public Const ID_BOW = "bow"
Public Const ID_ARROWS = "arrows"
Public Const ID_ARROWHEAD = "arrowhead"
Public Const ID_ARROWSHAFT = "arrowshaft"
Public Const TAG_NOT_SET = "Not Set"

'=============================================================
'                       Vulns Flags
'=============================================================

Enum eVulnFlags
    FL_SLASHING = 1
    FL_BLUDGEONING = 2
    FL_PIERCING = 4
    FL_FIRE = 8
    FL_COLD = 16
    FL_ACID = 32
    FL_LIGHTNING = 64
    FL_IMPERIL = 128
End Enum

'=============================================================
'                       Attack Height
'=============================================================
Enum eAttackHeight
    ATK_HIGH = 0
    ATK_MEDIUM = 1
    ATK_LOW = 2
End Enum

'=============================================================
'                       Logout Reasons
'=============================================================

Public Const LOGOUT_REASON_DIED = "Died"
Public Const LOGOUT_TIMER_DIED = 60 'logout 60 seconds after dying
Public Const LOGOUT_NO_HEALINGKIT = "No healing kits in inventory"


'=============================================================
'                       Colors
'=============================================================

Public Const COLOR_RED = 21
Public Const COLOR_GREEN = 0
Public Const COLOR_YELLOW = 4
Public Const COLOR_BRIGHT_YELLOW = 10
Public Const COLOR_WHITE = 2
Public Const COLOR_BLUE = 7
Public Const COLOR_PINK = 22
Public Const COLOR_CYAN = 13
Public Const COLOR_PURPLE = 5

'Chat Color Codes
Public Const CHAT_SYSTEM_MESSAGE = 0    'green text
Public Const CHAT_PUBLIC = 2            'White
Public Const CHAT_TELL = 4              '@tells
Public Const CHAT_SPELL_CASTING = 7     'all spell casting messages without spell words
Public Const CHAT_SPELL_RESULTS = 11    'more spells
Public Const CHAT_SPELL_WORDS = 17      'you say "Malar ...
Public Const CHAT_ALLEGIANCE = 18       'Allegiance chat
Public Const CHAT_MONSTER_ATTACKING_US = 21 'red, monster attacking us
Public Const CHAT_ATTACKING_MONSTER = 22 'red, us attacking monster
Public Const CHAT_GLOBAL_GENERAL = 27   'global GENERAL chat
Public Const CHAT_GLOBAL_TRADE = 28     'global TRADE chat
Public Const CHAT_GLOBAL_LFG = 29       'global LFG chat

'=============================================================
'                       Key and Mouse
'=============================================================

'Public Const KEYEVENTF_EXTENDEDKEY = &H1
'Public Const KEYEVENTF_KEYUP = &H2

'Public Const VK_UP = &H26
'Public Const VK_DOWN = &H28

'Public Const MOUSEEVENTF_ABSOLUTE = &H8000  '  absolute move
'Public Const MOUSEEVENTF_LEFTDOWN = &H2     '  left button down
'Public Const MOUSEEVENTF_LEFTUP = &H4       '  left button up
'Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '  middle button down
'Public Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
'Public Const MOUSEEVENTF_MOVE = &H1         '  mouse move
'Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '  right button down
'Public Const MOUSEEVENTF_RIGHTUP = &H10     '  right button up

