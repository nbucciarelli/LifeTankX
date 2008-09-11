VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILTER_DARKSIDE = "DarksideFilter.Filter"
Private Const FILTER_CHARSTATS = "DecalFilters.CharacterStats"
Private Const FILTER_NETECHO = "DecalFilters.EchoFilter2"
Private Const FILTER_WORLD = "DecalFilters.World"

Public WithEvents g_charFilter As DecalFilters.CharacterStats
Attribute g_charFilter.VB_VarHelpID = -1
Public WithEvents g_worldFilter As DecalFilters.World
Attribute g_worldFilter.VB_VarHelpID = -1

Private m_spellFilter As SpellFilter.Spells

Private WithEvents m_DSFilter As DarksideFilter.Filter
Attribute m_DSFilter.VB_VarHelpID = -1
Private WithEvents m_netEcho As DecalFilters.EchoFilter2
Attribute m_netEcho.VB_VarHelpID = -1


'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    Set g_charFilter = Nothing
    Set g_worldFilter = Nothing
    Set m_netEcho = Nothing
    Set m_DSFilter = Nothing
End Sub

Private Sub Class_Terminate()
    Set g_charFilter = Nothing
    Set g_worldFilter = Nothing
    Set m_netEcho = Nothing
    Set m_DSFilter = Nothing
    Set m_spellFilter = Nothing
End Sub

'#####################################################################################
'#
'#                                PROPERTIES
'#
'#####################################################################################

Public Property Get Server() As String
    If ValidFilter(m_DSFilter) Then Server = m_DSFilter.ServerName
End Property

Public Property Get ServerId() As eGameServer
    If ValidFilter(m_DSFilter) Then ServerId = m_DSFilter.ServerId
End Property

Public Property Get PlayerGUID() As Long
    If ValidFilter(m_DSFilter) Then PlayerGUID = m_DSFilter.Player.Guid
End Property

Public Property Get playerName() As String
    If ValidFilter(m_DSFilter) Then playerName = m_DSFilter.Player.Name
End Property

Public Property Get MaxHealth() As Long
    If ValidFilter(g_charFilter) Then MaxHealth = g_charFilter.EffectiveVital(eHealth)
End Property

Public Property Get MaxStam() As Long
    If ValidFilter(g_charFilter) Then MaxStam = g_charFilter.EffectiveVital(eStamina)
End Property

Public Property Get MaxMana() As Long
    If ValidFilter(g_charFilter) Then MaxMana = g_charFilter.EffectiveVital(eMana)
End Property

Public Property Get Health() As Long
    If ValidFilter(g_charFilter) Then Health = g_charFilter.Health
End Property

Public Property Get Stam() As Long
    If ValidFilter(g_charFilter) Then Stam = g_charFilter.Stamina
    End Property

Public Property Get Mana() As Long
    If ValidFilter(g_charFilter) Then Mana = g_charFilter.Mana
End Property

Public Property Get ActiveSpell(lNum As Long) As Enchantment
    If Valid(g_charFilter) Then
        If lNum < g_charFilter.EnchantmentCount Then
            Set ActiveSpell = g_charFilter.Enchantment(lNum)
        Else
            Set ActiveSpell = Nothing
        End If
    End If
End Property

Public Property Get ActiveSpellsCount() As Long
    If Valid(g_charFilter) Then ActiveSpellsCount = g_charFilter.EnchantmentCount
End Property

Public Property Get dsFilter() As DarksideFilter.Filter
    Set dsFilter = m_DSFilter
End Property

Public Function SpellLearned(ByVal SpellID As Long) As Boolean

    If Valid(g_charFilter) Then
        SpellLearned = (g_charFilter.SpellLearned(SpellID) <> 0)
    Else
        SpellLearned = False
    End If
    
End Function

'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

Private Function ValidFilter(ByVal objFilter As Variant) As Boolean
    ValidFilter = Valid(objFilter)
End Function


Private Function LoadNetecho() As Boolean
On Error GoTo InvalidFilter
    Dim bRet As Boolean
    
    Set m_netEcho = g_PluginSite.NetworkFilter(FILTER_NETECHO)
    bRet = Valid(m_netEcho)
    
Fin:
    LoadNetecho = bRet
    Exit Function
InvalidFilter:
    bRet = False
    Set m_netEcho = Nothing
    Resume Fin
End Function

Private Function LoadSpellFilter() As Boolean
On Error GoTo InvalidFilter
    
    Dim bRet As Boolean
    
    Set m_spellFilter = g_PluginSite.NetworkFilter("SpellFilter.Spells")
    
    bRet = Valid(m_spellFilter)
    
    If (bRet) Then
        Dim i As Long
        
        For i = 0 To 6000
            Dim spell As SpellFilter.spell
            Set spell = m_spellFilter.SpellByID(i)
            If Valid(spell) Then
                'MyDebug "spellfilter: spellid: " & spell.spellID & " name: " & spell.Name
                
                'MyDebug "<id=" & spell.SpellID & "><spl=" & spell.Name & ">"
                
                '        <spl=Spirit Drinker II><type=0><family=Spirit Drinker><sch=3><lvl=2><id=3254><icon=100673984><desc=Increases elemental damage bonus of a magic caster by 2%>
                '        <spl=Piercing Protection Self VI><type=10><family=Piercing Protection Self><sch=0><lvl=6><elemt=2><id=1138><icon=100668347><desc=Reduces damage the caster takes from Piercing by 60%.>
                '        <spl=Mana Conversion Mastery Self V><type=0><family=Mana Conversion Mastery Self><sch=2><lvl=5><id=657><icon=100668288><desc=Increases the caster's Mana Conversion skill by 30 points.>
                '        <spl=Flame Bolt IV><type=1><family=Flame Bolt><sch=1><lvl=4><elemt=3><id=83><icon=100668276><desc=Shoots a bolt of flame at the target.  The bolt does 31-60 points of fire damage to the first thing it hits.>
                
                MyDebug "<spl=" & spell.Name & "><type=" & spell.SpellType & "><family=" & spell.Name & "><sch=" & spell.SpellSchool & "><lvl=" & spell.Difficulty & "><id=" & spell.SpellID & "><icon=" & spell.Icon & "><desc=" & spell.Description & ">"

                'MyDebug "Spell Name: " & spell.Name
                'MyDebug "Spell Description: " & spell.Description
                'MyDebug "Spell Diff: " & spell.Difficulty
                'MyDebug "Spell Duration: " & spell.Duration
                'MyDebug "Spell Economy: " & spell.Economy
                'MyDebug "Spell Effect: " & spell.Effect
                'MyDebug "Spell EonCaster: " & spell.EffectOnCaster
                'MyDebug "Spell EonTarget: " & spell.EffectOnTarget
                'MyDebug "Spell Flags: " & spell.flags
                'MyDebug "Spell ID2: " & spell.ID2
                'MyDebug "Spell ManaCost: " & spell.ManaCost
                'MyDebug "Spell SortOrder: " & spell.SortOrder
                'MyDebug "Spell SpellID: " & spell.SpellID
                'MyDebug "Spell SpellSchool: " & spell.SpellSchool
                'MyDebug "Spell SpellType: " & spell.SpellType
                'MyDebug "Spell SpellVersion: " & spell.SpellVersion
                'MyDebug "Spell TargetMask: " & spell.TargetMask

                
                'SpellID,SpellID [Hex], Name              ,Icon [Hex],SortOrder,Difficulty,Duration,DurationUnk,Economy,Effect,EffectOnCaster,EffectOnTarget,Flags [Hex],ManaCost,RangeBase,RangeModifier,Researchable,Speed,SpellSchool,SpellType [Hex],SpellVersion,TargetMask [Hex],LUnknown2 [Hex],LUnknown3 [Hex],LUnknown6 [Hex], ID2,Description
                '1482,           0x5CA, Impenetrability II,0x60029BE,      3044,        50,    1800,-2.07525870,      1,   160,             0,           141,        0x4,      20,       25,            0,        True, 0.05,          3,            0x1,           1,             0x6,            0x0,     0xC4268000,            0x4,1482,"Improves a shield or piece of armor's armor value by 50 points. Target yourself to cast this spell on all of your equipped armor."
                '1483,           0x5CB,Impenetrability III,0x60029BE,      3046,       100,    1800,-2.07525870,      1,   160,             0,           141,        0x4,      30,       25,            0,        True,  0.1,          3,            0x1,           1,             0x6,            0x0,     0xC4268000,            0x6,1483,"Improves a shield or piece of armor's armor value by 75 points. Target yourself to cast this spell on all of your equipped armor."
                '2108,           0x83C, Brogard's Defiance,0x60029BE,      3053,       300,    3600,-2.07525870,      1,   160,             0,           141,       0x84,      70,       25,            0,       False, 0.15,          3,            0x1,           1,             0x6,            0x0,     0xC4268000,            0xE,2108,"Improves a shield or piece of armor's armor value by 220 points. Target yourself to cast this spell on all of your equipped armor."

                'MyDebug "spellfilter:spellid: " & spell.spellID & " name:" & spell.Name & "  effect: " & spell.Effect
            End If
        Next i
        
    End If
    
Fin:
    LoadSpellFilter = bRet
    Exit Function
InvalidFilter:
    bRet = False
    Set m_spellFilter = Nothing
    Resume Fin
End Function

Private Function LoadCharstats() As Boolean
On Error GoTo InvalidFilter
    Dim bRet As Boolean

    Set g_charFilter = g_PluginSite.NetworkFilter(FILTER_CHARSTATS)
    bRet = Valid(g_charFilter)

Fin:
    LoadCharstats = bRet
    Exit Function
InvalidFilter:
    bRet = False
    Set g_charFilter = Nothing
    Resume Fin
End Function

Private Function LoadWorldfilter() As Boolean
On Error GoTo InvalidFilter
    Dim bRet As Boolean

    Set g_worldFilter = g_PluginSite.NetworkFilter(FILTER_WORLD)
    bRet = Valid(g_worldFilter)

Fin:
    LoadWorldfilter = bRet
    Exit Function
InvalidFilter:
    bRet = False
    Set g_worldFilter = Nothing
    Resume Fin
End Function

Private Function LoadDarksideFilter() As Boolean
On Error GoTo InvalidFilter
    Dim bRet As Boolean
    
    Set m_DSFilter = g_PluginSite.NetworkFilter(FILTER_DARKSIDE)
    Set g_ds = m_DSFilter
    
    If Valid(m_DSFilter) Then
        Set g_Objects = m_DSFilter.GameObjects
        Set g_ACConst = m_DSFilter.ACConstants
        Call g_Events.SetDarksideFilter(g_ds)
    End If
    
    bRet = Valid(m_DSFilter)
   
Fin:
    LoadDarksideFilter = bRet
    Exit Function
InvalidFilter:
    bRet = False
    PrintErrorMessage "LoadDarksideFilter - " & Err.Description & " (line : " & Erl & ") - Err# " & Err.Number
    Set m_DSFilter = Nothing
    Resume Fin
End Function

Private Sub g_worldFilter_CreateObject(ByVal wObj As WorldObject)
On Error GoTo ErrorHandler

1    Dim objEntity As acObject

2    If Not g_Macro.Active Then Exit Sub
3    'If Not Valid(wObj) Then Exit Sub

4    Set objEntity = g_Objects.FindObject(wObj.Guid)

5    If Not Valid(objEntity) Then
        'Decal WorldFilter OnCreate caught the Object, but it doesn't exist in DS Filter!
6        MyDebug "g_worldFilter_CreateObject: does not exist in DSFilter"
7        MyDebug "g_worldFilter_CreateObject: Name: " & wObj.Name & " GUID: " & wObj.Guid
8    End If

Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "g_worldFilter_CreateObject - " & Err.Description & " (line : " & Erl & ") - Err# " & Err.Number
    Resume Fin
End Sub

'#####################################################################################
'#
'#                                  NETECHO
'#
'#####################################################################################
Private Sub m_netEcho_EchoMessage(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

'    Dim oNetMsg As NetMsgCls
'    Dim id As Long
'    Dim BytesArray() As Byte
'    Dim lLen As Long
'    Dim i As Long
'
'    'Research stuff
'    Select Case pMsg.Type
''        Case &H1E2&
''            MyDebug "Emote Text(" & pMsg.Type & ") - SenderId(DWORD):" & pMsg.Member("sender") & " - SenderName:" & pMsg.Member("senderName") & " - text:" & pMsg.Member("text")
''            Set oNetMsg = New NetMsgCls
''            oNetMsg.Data = pMsg.Data
''            'Call oNetMsg.DumpData
''
''            MyDebug "First DWORD: " & oNetMsg.ReadDWORD
''            MyDebug "Second DWORD: " & oNetMsg.ReadDWORD
''
''        Case &HF62C& 'Server text
''            MyDebug "ServerText - " & pMsg.Member("text")
''            Set oNetMsg = New NetMsgCls
''            oNetMsg.Data = pMsg.Data
''            'Call oNetMsg.DumpData
'
'        Case &HF7B0&
'            id = pMsg.Member("event")
'            Select Case id
'                Case &H38&   'direct tell
'                    MyDebug "- DirectTell Received -"
'                    'Set oNetMsg = New NetMsgCls
'                    'oNetMsg.Data = pMsg.Data
'                    'Call oNetMsg.DumpData
'
'                    BytesArray = pMsg.Data
'
''                    lLen = LenB(pMsg.Data)
''                    MyDebug "Len(pMsg.Data) = " & lLen
''                    ReDim BytesArray(0 To lLen - 1)
''                    Call CopyMemory(BytesArray(0), pMsg.Data, lLen)
'
'
'                    For i = LBound(BytesArray) To UBound(BytesArray)
'                        MyDebug "BytesArray(" & i & ") = " & BytesArray(i) & " --- chr: " & Chr(BytesArray(i))
'                    Next i
'
'                    'MyDebug "DirectTell - SenderID: " & pMsg.Member("source") & " - MagicNum: " & pMsg.Member("magic")
'            End Select
'
'    End Select
'
''    If Valid(g_Events) Then
''        Call g_Events.HandleNetworkMessages(pMsg)
''    End If
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "NetEcho - " & Err.Description
    Resume Fin
End Sub

'#####################################################################################
'#
'#                                  FRIENDS
'#
'#####################################################################################

Friend Function LoadFilters() As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    bRet = False
    
    'Load DarksideFilter
    If Not LoadDarksideFilter Then
        PrintErrorMessage "[Filters] DarksideFilter unavailable or not installed properly."
        GoTo Fin
    Else
        MyDebug "[Filters] DarksideFilter loaded."
    End If
    
    'Load netecho
    'If Not LoadNetecho Then
    '    PrintErrorMessage "[Filters] Echo Filter unavailable or not installed properly."
    '    GoTo Fin
    'Else
    '    MyDebug "[Filters] NetEcho loaded."
    'End If

    'Load CharStats
    If Not LoadCharstats Then
        PrintErrorMessage "[Filters] CharacterStats Filter unavailable or not installed properly."
        GoTo Fin
    Else
        MyDebug "[Filters] CharacterStats loaded."
    End If
    
    If Not LoadWorldfilter Then
        PrintErrorMessage "[Filters] World Filter unavailable or not installed properly"
        GoTo Fin
    Else
        MyDebug "[Filters] World Filter loaded."
    End If
  
    'Load SpellFilter
    If Not LoadSpellFilter Then
        PrintErrorMessage "[Filter] SpellFilter not loaded"
    Else
        MyDebug "[Filters] SpellFilter loaded."
    End If
    
    'Everything went fine
    bRet = True
    
Fin:
    LoadFilters = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsFilters.LoadFilters - " & Err.Description
    Resume Fin
End Function

'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################

Public Sub TestFilters()
On Error GoTo ErrorHandler
    
    MyDebug "----- Filters Test -------"
    MyDebug "Server : " & Server
    MyDebug "Player GUID : " & m_DSFilter.Player.Guid
    MyDebug "Player Name : " & m_DSFilter.Player.Name
    MyDebug "Monarch Name : " & m_DSFilter.Allegiance.Name
    MyDebug "Patron Name : " & m_DSFilter.Allegiance.Patron
    MyDebug "--------------------------"
    
    PrintMessage "Loading Player: " & m_DSFilter.Player.Name

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsFilters.TestFilters - " & Err.Description
    Resume Fin
End Sub




