VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsACEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'AC Events
Option Explicit

Private Const DEBUG_ME = False

Private WithEvents m_DSFilter As DarksideFilter.Filter
Attribute m_DSFilter.VB_VarHelpID = -1
Private WithEvents m_Objects As DarksideFilter.GameObjects
Attribute m_Objects.VB_VarHelpID = -1
Private WithEvents m_Items As DarksideFilter.acItems
Attribute m_Items.VB_VarHelpID = -1
Private WithEvents m_Fellowship As DarksideFilter.acFellowship
Attribute m_Fellowship.VB_VarHelpID = -1

Private m_bDeathMessageOccured As Boolean

Private Const TMR_CLOCK_CHECKER = 5 'check for system clock modification every 5 minutes
Private WithEvents m_tmrClockChecker As clsTimer
Attribute m_tmrClockChecker.VB_VarHelpID = -1

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    m_bDeathMessageOccured = False
    g_TotalKilled = 0
    g_TotalLooted = 0
    Set m_tmrClockChecker = CreateTimer
    Call SetDarksideFilter(Nothing)
End Sub

Private Sub Class_Terminate()
    Set m_tmrClockChecker = Nothing
    Call SetDarksideFilter(Nothing)
End Sub


'#####################################################################################
'#
'#                                  SETTERS
'#
'#####################################################################################

Friend Sub SetDarksideFilter(ByVal dsFilter As DarksideFilter.Filter)
On Error GoTo ErrorHandler

    Set m_DSFilter = dsFilter
    If Valid(m_DSFilter) Then
        Set m_Objects = m_DSFilter.GameObjects
        Set m_Items = m_DSFilter.GameObjects.Items
        Set m_Fellowship = m_DSFilter.GameObjects.Fellowship
    Else
        Set m_Objects = Nothing
        Set m_Items = Nothing
        Set m_Fellowship = Nothing
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.SetDarksideFilter"
    Resume Fin
End Sub

'######################################################################################
'#
'#                           CharacterStats EVENTS
'#
'######################################################################################

Private Sub g_charFilter_ActionComplete()
On Error GoTo ErrorHandler

    locDebug "clsACEvents.g_charFilter_ActionComplete"

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "g_charFilter_ActionComplete - " & Err.Description
    Resume Fin
End Sub


'######################################################################################
'#
'#                           DARKSIDE FILTER EVENTS
'#
'######################################################################################

Private Sub m_DSFilter_ForcedExit()
On Error GoTo ErrorHandler

    MyDebug "clsACEvents: ForcedExit()"
    
    'If Valid(g_Engine) Then
    '    Call g_Core.ForcedExit
    'End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.ForcedExit"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnMyDeath(ByVal sDeathMessage As String)
On Error GoTo ErrorHandler
  
    'SPK - Added death alert
    If g_ui.Options.chkAlertDeath.Checked Then
        Call PlaySound(SOUND_DEATH)
    End If
    
    If g_Macro.Ticking Then
        LogEvent "Macro died : " & sDeathMessage
        g_Macro.Died = True
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnMyDeath"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnRareLoot(ByVal aGuid As Long)
On Error GoTo ErrorHandler


Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnRareLoot"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnReady()
On Error GoTo ErrorHandler
    
    'OnReady seems to get called even when it's not true!
    'Exit Sub
    
    Dim bWasCasting As Boolean
    
    bWasCasting = False
    
    If g_Spells.Casting Then
        locDebug "DSFilter_OnReady: g_Spells.Casting is TRUE"
        bWasCasting = True
        Exit Sub
        'Call g_Spells.OnSpellCastComplete(, "OnReady")
    End If
    
    If g_Macro.Ticking Then
        locDebug "DSFilter_OnReady: g_Macro.Ticking"
        Call g_Macro.OnReady(True, bWasCasting)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnReady"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnActionFailure(ByVal iReason As Integer)
On Error GoTo ErrorHandler

    If g_Macro.Ticking Then
        If Not g_Macro.OnLastActionFailed(iReason) Then
            Select Case iReason
                Case FAIL_MISSILE_ATTACK_MISFIRED
                    MyDebug "Fail - Missile attack misfired"
                    ' JSC - FIXME - Increment MISSCOUNT?
                    
                Case FAIL_PROJECTILE_MISLAUNCHED
                    MyDebug "Fail - Projectile attack mislaunched"
                    ' JSC -- FIXME - BlackList them! -- or should I increment MISSCOUNT instead?
                    'Call g_Macro.Combat.Target.SetUserData(INT_BLIST_TIME, g_Core.Time + BLACKLIST_TIME)

                Case FAIL_SOLVED_QUEST_TOO_RECENTLY
                    PrintMessage "Failure - Quest solved too recently"
                
                Case FAIL_INVALID_COMMAND
                    Exit Sub
                    
                Case FAIL_NOT_ENOUGH_MANA_ON_ITEM
                    Exit Sub
                
                Case Else
                    MyDebug "GameEvents: Unkown error message #" & iReason & " [Hex=" & Hex(iReason) & "], reseting target & casting flag"
            End Select
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnActionFailure"
    Resume Fin
End Sub

Private Sub m_DSFilter_MeleeAttackComplete()
On Error GoTo ErrorHandler

    If g_Macro.Ticking Then
        locDebug "clsACEvents_DSFilter_MeleeAttackComplete"
        Call g_Macro.Combat.OnMeleeAttackComplete
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.MeleeAttackComplete"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnPlayerEvade(ByVal sAttacker As String)
On Error GoTo ErrorHandler

    If g_ui.Macro.chkEnableDOT.Checked Then
        Call g_DOT.playerEvade(sAttacker)
    End If

    ''This is here to find those pesky critters that are "invisible" to dsfilter
    'If g_Macro.isIdle And Not TargetScanner(g_Objects.Monsters) Then
    '    'Idle and No citters within range, yet we just took melee damage, WTF?
    '    ' Maybe we should not be shooting wars?
    '    If Not g_ui.Macro.chkNoWar.Checked Then
    '        'Ok, try select closest critter and see what that gets us
    '        Call g_Core.SendKey(g_Keys.KeyClosestMonster)
    '        MyDebug "clsACEvents.OnPlayerEvade: ST_IDLE, no target and evading melee damage"
    '        MyDebug "clsACEvents.OnPlayerEvade: LOST CRITTER!"
    '        Call g_Hooks.IDQueueAdd(g_Hooks.CurrentSelection)
    '        g_Macro.m_lostCritter = True
    '    End If
    'End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnPlayerEvade - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnTargetEvade(ByVal sTarget As String)
On Error GoTo ErrorHandler

    If g_ui.Macro.chkEnableDOT.Checked Then
        Call g_DOT.targetEvade(sTarget)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnPlayerEvade - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnInflictMeleeDamage(ByVal sTarget As String, ByVal lDamage As Long, ByVal bCrit As Boolean, ByVal sDamageType As String)
On Error GoTo ErrorHandler

    Call g_Macro.Combat.updateSecureTimer(3)

    If g_ui.Macro.chkEnableDOT.Checked Then
        Call g_DOT.giveMeleeDamage(sTarget, lDamage, bCrit, sDamageType)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnReceiveMeleeDamage - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnReceiveMeleeDamage(ByVal sAttacker As String, ByVal lDamage As Long, ByVal bCrit As Boolean, ByVal sDamageType As String, ByVal sLocation As String)
On Error GoTo ErrorHandler

    If g_Macro.Paused And g_Macro.adminPaused Then
        g_Macro.adminPaused = False
        Call TogglePause
    End If
    
    If g_ui.Macro.chkDangerZone.Checked Then
        Call g_Macro.Combat.tmrDanger.SetNextTime(3)    '3 seconds of danger time
    End If
    
    If g_ui.Macro.chkEnableDOT.Checked Then
        Call g_DOT.takeMeleeDamage(sAttacker, lDamage, bCrit, sDamageType, sLocation)
    End If
    
    'This is here to find those pesky critters that are "invisible" to decal/dsfilter
    'If g_Macro.isIdle And Not TargetScanner(g_Objects.Monsters) Then
    '    'Idle and No citters within range, yet we just took melee damage, WTF?
    '    ' Maybe we should not be shooting wars?
    '    If Not g_ui.Macro.chkNoWar.Checked Then
    '        'Ok, try select closest critter and see what that gets us
    '        Call g_Core.SendKey(g_Keys.KeyClosestMonster)
    '        MyDebug "clsACEvents.OnReceiveMeleeDamage: ST_IDLE, no target and taking melee damage!"
    '        MyDebug "clsACEvents.OnReceiveMeleeDamage: LOST CRITTER! IDing now"
    '        Call g_Hooks.IDQueueAdd(g_Hooks.CurrentSelection)
    '        g_Macro.m_lostCritter = True
    '    End If
    'End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnReceiveMeleeDamage - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnReceiveDeathMessage(ByVal DeathMessage As String)
On Error GoTo ErrorHandler

    'for death message filter
    m_bDeathMessageOccured = True
    g_TotalKilled = g_TotalKilled + 1

    g_Macro.m_lostCritter = False

    If g_ui.Loot.chkEnableLooting.Checked Then
        Call g_Macro.incLoot
    End If

    If g_Spells.Casting Or Not (g_Macro.OkToCast) Then
        Call g_Spells.ResetCastingFlag("Death Message")
    End If
    
'    If g_Macro.Ticking Then
'        Call g_Macro.OnMonsterDeathMessage(DeathMessage)
'    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnReceiveDeathMessage - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnReceiveTell(ByVal sMessage As String, ByVal sSenderName As String, ByVal lSenderGUID As Long, ByVal vData As Variant)
On Error GoTo ErrorHandler
    
    'Log tells
    If g_ui.MsgLog.chkMonitorTells.Checked Then
        Call g_ui.MsgLog.AddLog(TYPE_TELL, sSenderName, sMessage)
    End If
        
    If g_ui.Options.chkAlertTell.Checked Then
        Dim playTell As Boolean
        Dim objPlayer As acObject
        Dim objEntity As acObject
        playTell = True
        
        Set objPlayer = g_Objects.Player
        Set objEntity = g_Objects.FindObject(lSenderGUID)
        If Valid(objEntity) Then
            If (objEntity.ObjectType = TYPE_MERCHANT) Or (objEntity.ObjectType = TYPE_MONSTER) Then playTell = False
        End If
        
        If playTell Then
                Call PlaySound(SOUND_TELL)
        Else
                MyDebug "No PLAYTELL, Monster Talking " & lSenderGUID
                ' do nothing
        End If
    End If

    'Only execute when macro is running
    If g_Macro.Ticking Then
    
        'Watch for admin
        Call g_AntiBan.CheckForAdminTell(sMessage, sSenderName, lSenderGUID, vData)
    
        'Remote command
        Call g_RemoteCmd.HandleRemoteCommands(sSenderName, sMessage)
        
        If g_ui.Irc.ConnectedToChannel And g_ui.AntiBan.chkReportTells.Checked Then
            Call g_ui.Irc.SendChanMessage("[@Tell] " & sSenderName & " tells you, """ & sMessage & """")
        End If
    
        'Fellow Auto-List commands
        If (g_FellowList.Count > 0) And g_ui.Macro.chkAutoFellowList.Checked Then
            Call g_Macro.FellowCmd.CheckCommand(sSenderName, sMessage)
        End If
        
        'Fellowship Command - Password Check
        If SameText(sMessage, g_ui.Macro.txtFellowPassword.Text) And g_ui.Macro.chkEnableAutoFellow.Checked Then
            Call g_Macro.FellowCmd.OnPasswordAccepted(sSenderName, lSenderGUID)
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnReceiveTell"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnLocalChat(ByVal sMessage As String, ByVal sSenderName As String, ByVal lSenderID, ByVal lChatType As Long)
On Error GoTo ErrorHandler
    
    '0x00    Broadcast (e.g. allegiance MOTD)
    '0x02    Public Chat
    '0x03    Private Tell
    '0x04    Outgoing Tell (e.g. 'You tell ...')
    '0x07    Magic Spell Results
    '0x0c    NPC Chat
    '0x11    Player Spellcasting
    '0x12    Creature Chat (e.g. 'Fellow warriors, aid me!')
    '0x17    Recall (e.g. 'Player is recalling home.')

    If (lSenderID = m_Objects.Player.Guid) Then
        ' Exit if this player is talking
        Exit Sub
    End If
    
    'MyDebug "OnLocalChat (" & Hex(lChatType) & "): " & sSenderName & " : " & sMessage
    
    If (lChatType And &H2&) Then
    
        'MyDebug "OnLocalChat (" & Hex(lChatType) & "): " & sSenderName & " : " & sMessage
        
        'Log open chat messages mentioning our name
        If g_ui.MsgLog.chkMonitorPublic.Checked And InStr(LCase(sMessage), LCase(g_Objects.Player.Name)) Then
            Call g_ui.MsgLog.AddLog(TYPE_PUBLIC_CHAT, sSenderName, sMessage)
        End If
    
        If g_ui.Options.chkAlertOpenChat.Checked Then
            Call PlaySound(SOUND_OPENCHAT)
        End If
    
        If g_Macro.Ticking And g_ui.Irc.ConnectedToChannel And g_ui.AntiBan.chkReportOpenChat.Checked Then
            Call g_ui.Irc.SendChanMessage("[Open Chat] " & sSenderName & " says, """ & sMessage & """")
        End If
    
    End If
    
    'If Not (lChatType = &H12&) And Not (lChatType = &HC&) And Not (lChatType = &H17&) Then
    '    Call g_Events.HandleConsoleMessage(sMessage, lChatType)
    'End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnLocalChat - " & Err.Description
    Resume Fin
End Sub

'Private Sub m_DSFilter_OnServerChat(ByVal sMessage As String, ByVal lType As Long)
'On Error GoTo ErrorHandler
'
'    '0x00    Broadcast (e.g. allegiance MOTD)
'    '0x02    Public Chat
'    '0x03    Private Tell
'    '0x04    Outgoing Tell (e.g. 'You tell ...')
'    '0x07    Magic Spell Results
'    '0x0c    NPC Chat
'    '0x11    Player Spellcasting
'    '0x12    Creature Chat (e.g. 'Fellow warriors, aid me!')
'    '0x17    Recall (e.g. 'Player is recalling home.')
'
'    'MyDebug "OnServerChat (" & Hex(lType) & "): " & sMessage
'
'    'If Not (lType = &H12&) And Not (lType = &HC&) And Not (lType = &H7&) And Not (lType = &H11&) And Not (lType = &H17&) Then
'    '    Call g_Events.HandleConsoleMessage(sMessage, lType)
'    'End If
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "ACEvents.OnServerChat - " & Err.Description
'    Resume Fin
'End Sub

'OnAddEnchantment
Private Sub m_DSFilter_OnAddEnchantment(ByVal SpellID As Long, ByVal lFamily As Long, ByVal dStartTime As Double, ByVal dTimeElapsed As Double, ByVal dDuration As Double, ByVal lSourceGUID As Long, ByVal iLayer As Long)
On Error GoTo ErrorHandler
    Dim objSpell As clsSpell

    MyDebug "OnAddEnchantment: " & SpellID & " :lFamily: " & lFamily & " :dDuration: " & dDuration & " :layer: " & iLayer

    'Set objSpell = g_Spells.sBuffs.FindSpellByID(spellID)
    'If Valid(objSpell) Then
    '    'Call g_Buffer.CheckCastedSpellObj(objSpell)
    'End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnAddEnchantment - " & Err.Description
    Resume Fin
End Sub

'OnRemoveEnchantment
Private Sub m_DSFilter_OnRemoveEnchantment(ByVal SpellID As Long, ByVal iLayer As Long)
On Error GoTo ErrorHandler
    Dim objSpell As clsSpell
    
    MyDebug "OnRemoveEnchantment: " & SpellID & " :layer: " & iLayer
    
    'Set objSpell = g_Spells.Items.FindSpellByID(spellID)
    'If Valid(objSpell) Then
    '    Call g_Buffer.CheckExpiredItemSpellObj(spellID)
    'End If
    Set objSpell = g_Spells.sBuffs.FindSpellByID(SpellID)
    If Valid(objSpell) Then
        Call g_Buffer.CheckExpiredSpell(objSpell.SpellName)
    End If

Fin:
    Set objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnRemoveEnchantment - " & Err.Description
    Resume Fin
End Sub

'OnRemoveMultEnchantment
Private Sub m_DSFilter_OnRemoveMultEnchantment(ByVal SpellID As Long, ByVal iLayer As Long)
On Error GoTo ErrorHandler
    Dim objSpell As clsSpell
    
    MyDebug "OnRemoveMultEnchantment: " & SpellID & " :layer: " & iLayer
    
    'Set objSpell = g_Spells.Items.FindSpellByID(spellID)
    'If Valid(objSpell) Then
    '    Call g_Buffer.CheckExpiredItemSpellObj(spellID)
    'End If

    Set objSpell = g_Spells.sBuffs.FindSpellByID(SpellID)
    If Valid(objSpell) Then
        Call g_Buffer.CheckExpiredSpell(objSpell.SpellName)
    End If

Fin:
    Set objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnRemoveMultEnchantment - " & Err.Description
    Resume Fin
End Sub

'OnSilentRemoveEnchantment
Private Sub m_DSFilter_OnSilentRemoveEnchantment(ByVal SpellID As Long, ByVal iLayer As Long)
On Error GoTo ErrorHandler
    Dim objSpell As clsSpell
    
    MyDebug "OnSilentRemoveEnchantment: " & SpellID & " :layer: " & iLayer
    
    'Set objSpell = g_Spells.Items.FindSpellByID(spellID)
    'If Valid(objSpell) Then
    '    Call g_Buffer.CheckExpiredItemSpellObj(spellID)
    'End If

    Set objSpell = g_Spells.sBuffs.FindSpellByID(SpellID)
    If Valid(objSpell) Then
        Call g_Buffer.CheckExpiredSpell(objSpell.SpellName)
    End If

Fin:
    Set objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnSilentRemoveEnchantment - " & Err.Description
    Resume Fin
End Sub

'OnSilentRemoveMultEnchantment
Private Sub m_DSFilter_OnSilentRemoveMultEnchantment(ByVal SpellID As Long, ByVal iLayer As Long)
On Error GoTo ErrorHandler
    Dim objSpell As clsSpell
    
    MyDebug "OnSilentRemoveMultEnchantment: " & SpellID & " :layer: " & iLayer
    
    'Set objSpell = g_Spells.Items.FindSpellByID(spellID)
    'If Valid(objSpell) Then
    '    Call g_Buffer.CheckExpiredItemSpellObj(spellID)
    'End If
    
    Set objSpell = g_Spells.sBuffs.FindSpellByID(SpellID)
    If Valid(objSpell) Then
        Call g_Buffer.CheckExpiredSpell(objSpell.SpellName)
    End If

Fin:
    Set objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnSilentRemoveMultEnchantment - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnSpellVisible(ByVal SpellName As String)
On Error GoTo ErrorHandler

    locDebug "OnSpellVisible name: " & SpellName
    
    If (g_Macro.State = ST_REBUFF) Or (g_Macro.State = ST_BUDDYREBUFF) Or (g_Macro.State = ST_FELLOWHEAL) Then
        'Do nothing
    Else
        locDebug "-- DSFilter_OnSpellVisible: " & SpellName
        Call g_Spells.OnSpellCastComplete(True, "OnSpellVisible")
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnSpellVisible - " & Err.Description
    Resume Fin
End Sub

Private Sub m_DSFilter_OnSpellCast(ByVal SpellID As Long)
On Error GoTo ErrorHandler

    locDebug "OUTGOING: OnSpellCast event: spellID: " & SpellID
    
    ' Set spell casting start flags
    Call g_Spells.OnSpellCastBegin("OnSpellCast Event")
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnSpellCast"
    Resume Fin
End Sub

Private Sub m_DSFilter_OnSpellCastOn(ByVal SpellID As Long, ByVal objTarget As acObject)
On Error GoTo ErrorHandler

    locDebug "OUTGOING: OnSpellCastOn event: spellId: " & SpellID & "  target: " & objTarget.Name
    
    ' The problem is that this get's called as soon as the toon STARTS to cast the spell (war/vuln/yield spells)
    
    ' Set spell casting start flags
    Call g_Spells.OnSpellCastBegin("OnSpellCastOn Event")

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnSpellCastOn"
    Resume Fin
End Sub


'Private Sub m_DSFilter_OnLoginComplete()
'On Error GoTo ErrorHandler
'        'Call g_Engine.OnLoginComplete
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "ACEvents.OnLoginComplete - " & Err.Description
'    Resume Fin
'End Sub

'Private Sub m_DSFilter_OnLogin()
'On Error GoTo ErrorHandler
'        'Call g_Engine.OnLogin
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "ACEvents.OnLoginComplete - " & Err.Description
'    Resume Fin
'End Sub


'######################################################################################
'#
'#                                 GENERAL OBJECT EVENTS
'#
'######################################################################################

Private Sub m_Objects_OnCreateObject(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    If Not Valid(obj) Then Exit Sub
    
    locDebug "clsACEvents.OnCreateObject: " & obj.Name & " : " & obj.Guid
    
1    Select Case obj.ObjectType
        Case TYPE_PLAYER
            'Check if this player is an admin or unfriendly
            If obj.Guid <> m_Objects.Player.Guid Then
                Call initObjectExtra(obj)
                Call g_AntiBan.CheckAlerts(obj)
            End If
            
2        Case TYPE_MONSTER
            Call initObjectExtra(obj)
            Call InitMonster(obj)
            
            If g_ui.Options.chkDetectList.Checked Then
                If g_ui.Options.NameInDetectList(obj.Name, True) Then
                    PrintMessage "Detected: " & obj.Name & " [" & obj.Loc.Coords & "]", COLOR_PURPLE
                End If
            End If

3        Case TYPE_ITEM
            If Not (obj.itemType = ITEM_UNKNOWN) Then
                Call initObjectExtra(obj)
                Call g_Macro.Loot.OnCreateObject(obj)
            End If
            
            If g_Macro.State = ST_BUYSELL Then
                Call initObjectExtra(obj)
                Call PhatLoot.vendorEvent("OnCreateObject", obj)
            End If
            
            If g_ui.Options.chkDetectList.Checked Then
                If g_ui.Options.NameInDetectList(obj.Name, False) Then
                    PrintMessage "Detected: " & obj.Name & " [" & obj.Loc.Coords & "]", COLOR_PURPLE
                End If
            End If

    End Select

    'Tell the macro something has just changed in the environment
    If g_Macro.Ticking And g_Macro.State = ST_IDLE Then
        Call g_Macro.ResetTargetScanTimer
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnCreateObject - " & Err.Description & " - " & Erl
    Resume Fin
End Sub

Public Sub initObjectExtra(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler
    
    If obj.SetUserDataCount(NUM_EXTRA_OBJECT_INFO) Then
        Call obj.SetUserData(B_ENABLED, True)
        Call obj.SetUserData(B_CAN_BE_IMPERILED, True)
        Call obj.SetUserData(B_CAN_BE_VULNED, True)
        Call obj.SetUserData(B_CAN_BE_YIELDED, True)
        Call obj.SetUserData(INT_MISSCOUNT, 0)
        Call obj.SetUserData(INT_BLISTCOUNT, 0)
        Call obj.SetUserData(INT_SALVAGECOUNT, 0)
        Call obj.SetUserData(INT_YIELD_TRYS_LEFT, 3)
        Call obj.SetUserData(B_DANGEROUS, False)
        Call obj.SetUserData(B_WAR_CHECK, False)
        Call obj.SetUserData(INT_BLIST_TIME, False)
        
        Call obj.SetUserData(B_LOOTED, False)
        Call obj.SetUserData(B_MACO_PICKUP, False)
    Else
        PrintErrorMessage "clsACEvents.initObjectExtra - Unable to allocate extra object info"
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.initObjectExtra: " & Err.Description & " - " & Erl
    Resume Fin
End Sub

Private Sub m_Objects_OnIdentifyObject(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler
    
    If Not Valid(obj) Then Exit Sub
    
    'Update the loot id Queue if needed
    Call g_Macro.Loot.OnIdentifyObject(obj)
    
    'Update the salvager Id queue
    Call g_Macro.Salvager.OnIdentifyObject(obj)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnIdentifyObject"
    Resume Fin
End Sub

Private Sub m_Objects_OnApproachVendor(ByVal obj As DarksideFilter.acObject, ByVal vendorInv As DarksideFilter.colObjects)
On Error GoTo ErrorHandler
    
    If Not Valid(obj) Then Exit Sub
    
    ' Setup the Vendor Object and inventory
    locDebug "OnApproachVendor: Vendor:" & obj.Name
    
    Call PhatLoot.setVendor(obj)
    Call PhatLoot.setVendorInv(vendorInv)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnApproachVendor"
    Resume Fin
End Sub

Private Sub m_Objects_OnRareDetected(ByVal obj As DarksideFilter.acObject)
On Error Resume Next
    If g_ui.Options.chkFilterIdMsg.Checked Or g_Macro.Loot.silentID Then Exit Sub
     If ((g_Hooks.CurrentSelection = obj.Guid) Or (obj.Wielder <> 0)) Then
        PrintMessage "*** [Rare Detected] ***  " & obj.Name
        'Call PlaySound(SOUND_RARE)
    End If
End Sub

Private Sub m_Objects_OnEpicDetected(ByVal obj As DarksideFilter.acObject, ByVal sSpellName As String)
On Error Resume Next
    If g_ui.Options.chkFilterIdMsg.Checked Or g_Macro.Loot.silentID Then Exit Sub
    If ((g_Hooks.CurrentSelection = obj.Guid) Or (obj.Wielder <> 0)) And g_ui.Loot.chkPickupMajors.Checked Then PrintMessage "[Epic Detected] " & obj.Name & " : " & sSpellName
End Sub

Private Sub m_Objects_OnMajorDetected(ByVal obj As DarksideFilter.acObject, ByVal sSpellName As String)
On Error Resume Next
    If g_ui.Options.chkFilterIdMsg.Checked Or g_Macro.Loot.silentID Then Exit Sub
    If ((g_Hooks.CurrentSelection = obj.Guid) Or (obj.Wielder <> 0)) And g_ui.Loot.chkPickupMajors.Checked Then PrintMessage "[Major Detected] " & obj.Name & " : " & sSpellName
End Sub

Private Sub m_Objects_OnMinorDetected(ByVal obj As DarksideFilter.acObject, ByVal sSpellName As String)
On Error Resume Next
    If g_ui.Options.chkFilterIdMsg.Checked Or g_Macro.Loot.silentID Then Exit Sub
    If ((g_Hooks.CurrentSelection = obj.Guid) Or (obj.Wielder <> 0)) And g_ui.Loot.chkPickupMinors.Checked Then PrintMessage "[Minor Detected] " & obj.Name & " : " & sSpellName
End Sub

Private Sub m_Objects_OnObjectDeath(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler
 
    locDebug "clsACEvent.OnObjectDeath: " & obj.Name & " : " & obj.Guid
 
    g_Macro.m_lostCritter = False
    
    'Always call this, even when macro is not "runing"
    Call g_Macro.OnEntityDeath(obj)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnObjectDeath"
    Resume Fin
End Sub

Private Sub m_Objects_OnRemoveObject(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    If Not Valid(obj) Then Exit Sub
 
    locDebug "clsACEvent.OnRemoveObject: " & obj.Name & " : " & obj.Guid
    
    'Always call this, even when macro is not "running"
    Call g_Macro.OnEntityNoLongerAvailable(obj)
    
    If g_Macro.State = ST_SALVAGE Then
        Call g_Macro.Salvager.OnRemoveObject(obj)
    End If
    
    If g_Macro.State = ST_BUYSELL Then
        Call PhatLoot.vendorEvent("OnRemoveObject")
    End If
    
    If obj.itemType = ITEM_CORPSE Then
        Call g_Macro.Loot.OnIdentifyObject(obj)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnRemoveObject"
    Resume Fin
End Sub

Private Sub m_Objects_OnObjectMoved(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    If Not Valid(obj) Then Exit Sub
    
    If g_Nav.NavType = NAVTYPE_FOLLOW And Valid(g_Nav.objToFollow) Then
        If obj.Guid = g_Nav.objToFollow.Guid Then
            Call g_Nav.OnTargetMoved
        End If
    End If
    
    'If (obj.ObjectType = TYPE_MONSTER) And Not (landblockInRange(obj.Loc.landblock)) Then
    '    MyDebug "clsACEvents.OnObjectMoved: out of range: " & obj.Name & " : " & obj.Guid
    '
    '    'If it's a monster and it has a different landblock than this player, get rid of it in 30 seconds
    '    obj.timeData = g_ds.Time + 30
    'End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnObjectMoved"
    Resume Fin
End Sub

Private Sub m_Objects_OnSelectObject(ByVal obj As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    If Not Valid(obj) Then Exit Sub

    'If we are in Melee combat mode and have a target, make sure we have it selected
    If IsMelee And Valid(g_Macro.Combat.Target) Then
        'make sure target is selected
        If g_Hooks.CurrentSelection <> g_Macro.Combat.Target.Guid Then
            locDebug "Target is not selected - Forcing selection."
            Call g_Service.SelectObject(g_Macro.Combat.Target)
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnSelectObject"
    Resume Fin
End Sub


'######################################################################################
'#
'#                              ITEMS/INVENTORY EVENTS
'#
'######################################################################################


Private Sub m_Items_OnOpenContainer(ByVal objContainer As DarksideFilter.acObject, ByVal colItems As DarksideFilter.colObjects)
On Error GoTo ErrorHandler
    
    'if we're not in loot mode, don't even bother
    If (g_Macro.State = ST_LOOT) Then
        Call g_Macro.Loot.OnOpenContainer(objContainer, colItems)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnOpenContainer"
    Resume Fin
End Sub

Private Sub m_Items_OnStackSizeChanged(ByVal objItem As DarksideFilter.acObject)
On Error GoTo ErrorHandler
    
    If g_Macro.State = ST_LOOT Then
        Call g_Macro.Loot.CheckPickup(objItem.Guid, True)
    End If
    
    If g_Macro.State = ST_BUYSELL Then
        Call PhatLoot.vendorEvent("OnStackSizeChanged")
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnStackSizeChanged"
    Resume Fin
End Sub

Private Sub m_Items_OnEquipItem(ByVal objItem As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    MyDebug "m_Items_OnEquipItem event: " & objItem.Name
    'RaiseEvent g_Macro.m_Equip.OnItemEquipped

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnEquipItem"
    Resume Fin
End Sub

Private Sub m_Items_OnReceiveItem(ByVal objItem As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    If g_Macro.State = ST_BUYSELL Then
        Call PhatLoot.vendorEvent("OnReceiveItem")
    End If

    'Make sure we update the Loot list so items make it onto salvage list (if needed)
    Call g_Macro.Loot.OnItemPickup(objItem)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnReceiveItem"
    Resume Fin
End Sub

Private Sub m_Items_OnGiveItem(ByVal objItem As DarksideFilter.acObject, ByVal objDest As DarksideFilter.acObject)
On Error GoTo ErrorHandler

    If g_Macro.State = ST_BUYSELL Then
        Call PhatLoot.vendorEvent("OnGiveItem")
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACEvents.OnGiveItem"
    Resume Fin
End Sub

'######################################################################################
'#
'#                           FELLOWSHIP EVENTS
'#
'######################################################################################

Private Sub m_DSFilter_OnFellowshipMessage(ByVal sSenderName As String, ByVal sMsg As String)
On Error GoTo ErrorHandler

    If (g_FellowList.Count > 0) And InStr(sMsg, "!list") Then
         SendFellowshipMessage g_FellowList.getListString
    End If

    If Not (sSenderName = m_DSFilter.Player.Name) Then
        If g_FellowList.crashState And InStr(sMsg, "!crash") Then
            Call g_FellowList.foundCrash
        End If
        If g_ui.MsgLog.chkMonitorFellowship.Checked Then
            Call g_ui.MsgLog.AddLog(TYPE_FELLOWSHIP, sSenderName, sMsg)
        End If
        
        'If g_Core.InitComplete And Valid(g_RemoteCmd) Then
        '    Call g_RemoteCmd.RemoteRedirectChatToIRC("[Fellowship]" & sMsg)
        'End If
        
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ACEvents.OnFellowshipMessage - " & Err.Description
    Resume Fin
End Sub

Private Sub m_Fellowship_OnQuit(ByVal objFellow As acObject)
On Error Resume Next
    
    If g_Objects.IsSelf(objFellow) Then
        locDebug "You have left the fellowship: " & g_Objects.Fellowship.Name
        Call g_FellowList.Clear
    Else
        MyDebug objFellow.Name & " has left your fellowship: " & g_Objects.Fellowship.Name
        ' If there are people on the waiting list, send out a crash query message
        If (g_FellowList.Count > 0) Then
            Call g_FellowList.startCrashTest(objFellow.Name)
        End If
    End If
End Sub

Private Sub m_Fellowship_OnRecruit(ByVal objFellow As acObject)
On Error Resume Next
    
    MyDebug objFellow.Name & " joined the fellowship"
    
    If (g_FellowList.Count > 0) And g_FellowList.checkOnList(objFellow.Name) Then
        Call g_FellowList.removeFromLine(objFellow.Name)
    End If
End Sub

Private Sub m_Fellowship_OnJoin()
On Error Resume Next
    'join
    MyDebug "You joined fellowship: " & m_Fellowship.Name
End Sub

Private Sub m_Fellowship_OnUpdateFellowStats(ByVal objFellow As acObject)
On Error Resume Next
    
    'We have received an update on the stats of a fellow member
    Dim oFellow As acObject
    
    If g_ui.Macro.chkHealFellow.Checked Or g_ui.Macro.chkStamFellow.Checked Then
    
        For Each oFellow In g_Objects.Fellowship
            If objFellow.Name = oFellow.Name Then
                'MyDebug "Updating stats for fellow member: " & objFellow.Name
                oFellow.Health = objFellow.Health
                oFellow.Stamina = objFellow.Stamina
                
                If (objFellow.Health <= GetPercent(objFellow.MaxHealth, g_Data.FellowHealthThreshold)) Then
                    'Set global FellowNeedsHealing flag
                    locDebug "OnUpdateFellowStats: needs healing: " & objFellow.Name
                    Call g_Macro.SetFellowNeedsHealing(True)
                End If
                If (objFellow.Stamina <= GetPercent(objFellow.MaxStamina, g_Data.FellowStamThreshold)) Then
                    'Set global FellowNeedsStam flag
                    locDebug "OnUpdateFellowStats: needs Staming: " & objFellow.Name
                    Call g_Macro.SetFellowNeedsStam(True)
                End If
            End If
        Next oFellow
        
    End If
End Sub

Private Sub m_Fellowship_OnCreate()
On Error Resume Next
    MyDebug "Created a new fellowship"
End Sub

Private Sub m_Fellowship_OnKick(ByVal objFellow As acObject)
On Error Resume Next
        ' If there are people on the waiting list, notify the next in line
        If (g_FellowList.Count > 0) Then
            Call g_FellowList.notifyNext
        End If
End Sub

'Private Sub m_Fellowship_OnDisband()
'On Error Resume Next
'    'disband
'    LogEvent "The fellowship has been disbanded."
'
'    If g_ui.Options.chkAlertDisband.Checked Then
'        Call PlaySound(SOUND_FELLOW_DISBAND)
'    End If
'End Sub
'
'Private Sub m_Fellowship_OnRecvInvite()
'    MyDebug "Receiving fellowship invitation"
'End Sub

'######################################################################################


'missile Attack hit environment
Private Sub OnHitEnvironment()
On Error Resume Next
    If Valid(g_Macro.Combat.Target) And (g_Macro.State = ST_COMBAT_ATTACKING) Then
        ' JSC - increment the MISSCOUNT
        Dim missCount As Integer
        missCount = g_Macro.Combat.Target.UserData(INT_MISSCOUNT) + 1
        Call g_Macro.Combat.Target.SetUserData(INT_MISSCOUNT, missCount)
        'MyDebug "clsACEvents.HandleConsoleMessage: +1 MISSCOUNT: " & missCount & " for: " & g_Macro.Combat.Target.Name
    End If

End Sub



'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################

'HandleStatusMessage
Public Function HandleStatusMessage(ByVal bstrMsg As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim bFilterMsg As Boolean
    
    If InStr(bstrMsg, "too busy") Or InStr(bstrMsg, "Casting") Then
        bFilterMsg = True
    ElseIf InStr(bstrMsg, "is already a member of a fellowship") Then
        'Need to let them know they are aleady in a fellow and they need to drop the old one
    Else
        bFilterMsg = False
    End If

Fin:
    HandleStatusMessage = bFilterMsg
    Exit Function
ErrorHandler:
    PrintErrorMessage "HandleStatusMessage - " & Err.Description
    Resume Fin
End Function


Public Function HandleConsoleMessage(ByVal bstrMsg As String, ByVal pColor As Long) As Boolean
On Error GoTo ErrorHandler
    
    Dim bFilterMsg As Boolean
    Dim Msg As String
    Dim ircString As String
    bFilterMsg = False
    
    bstrMsg = Trim(bstrMsg)
    
    If Not (pColor = 27) And Not (pColor = 28) And Not (pColor = 29) Then
        locDebug "clsACEvents.HandleConsoleMessage:(" & pColor & ") " & bstrMsg
    End If
    
    If InStr(bstrMsg, "Tell:IIDString") Then
        'New click to talk looks like: [Fellowship] <Tell:IIDString:1343014699:ToonName>ToonName<\Tell> says, "Shurov Thiril"
        'So need to get rid of all the extra crap and just keep: ToonName says, "Shurov Thiril"
        Dim iLen As Integer, aPos As Integer, bPos As Integer
        Dim idString As String, aString As String, bString As String
        idString = Replace(bstrMsg, "<\Tell>", "")
        'MyDebug "HandleConsoleMessage:Replace: idString: " & idString
        aPos = InStr(idString, "<")
        bPos = InStr(idString, ">")
        iLen = Len(idString)
        If (aPos <= 0) Then aPos = 1
        aString = Left(idString, aPos - 1)
        bString = Right(idString, (iLen - bPos))
        idString = aString & bString
        'MyDebug "HandleConsoleMessage:after: idString: " & idString
        bstrMsg = idString
    End If
    
    'AntiBan feature - Check to see if there's an admin talking
    Call g_AntiBan.CheckConsoleForAdmin(bstrMsg, pColor)
    
    '=====================================
    ' Chat Log
    '=====================================
    If pColor = 2 Or _
        pColor = 3 Or _
        pColor = 4 Or _
        pColor = 8 Or _
        pColor = 9 Or _
        pColor = 10 Or _
        pColor = 12 Or _
        pColor = 19 Then
            Call LogChatMessage(Left(bstrMsg, Len(bstrMsg) - 1)) 'get rid of chr(10)
        'pColor = 22
    End If
    
    '=====================================
    ' Chat Console Filters
    '=====================================
    If (g_ui.Options.chkFilterSpellcasting.Checked And (pColor = CHAT_SPELL_CASTING Or pColor = CHAT_SPELL_WORDS Or pColor = CHAT_SPELL_RESULTS)) Then
        bFilterMsg = True   'tell the client we dont want to display this message in the chat console
    ElseIf (pColor = CHAT_MONSTER_ATTACKING_US Or pColor = CHAT_ATTACKING_MONSTER) Then
        ' Seperate options for Evades and melee damage
        If g_ui.Options.chkFilterMeleeEvade.Checked And _
        (InStr(bstrMsg, "You evaded") Or InStr(bstrMsg, "evaded your attack")) Then
            bFilterMsg = True
        ElseIf g_ui.Options.chkFilterMelee.Checked And (InStr(bstrMsg, "points of") Or InStr(bstrMsg, "point of")) Then
            bFilterMsg = True
        End If
    ElseIf (g_ui.Options.chkFilterDeathMsg.Checked And pColor = CHAT_SYSTEM_MESSAGE) And m_bDeathMessageOccured Then
        If InStr(bstrMsg, "Your task is complete") Then
            bFilterMsg = False
        ElseIf (InStrRev(bstrMsg, "!") = (Len(bstrMsg) - 1)) And Not (InStr(bstrMsg, "fellow")) Then
            'MyDebug "clsACEvents: filtered death message: (" & InStrRev(bstrMsg, "!") & ") " & bstrMsg
            bFilterMsg = True
            m_bDeathMessageOccured = False
        End If
    ElseIf (pColor = CHAT_GLOBAL_GENERAL) And g_ui.Options.chkFilterGlobalChat.Checked Then
        'global [General] channel
        bFilterMsg = True
    ElseIf (pColor = CHAT_GLOBAL_TRADE) And g_ui.Options.chkFilterGlobalTrade.Checked Then
        'global [Trade] channel
        bFilterMsg = True
    ElseIf (pColor = CHAT_GLOBAL_LFG) And g_ui.Options.chkFilterGlobalLFG.Checked Then
        'global [LFG] channel
        bFilterMsg = True
    ElseIf (pColor = CHAT_ALLEGIANCE) Then
        'Allegiance chat
    End If

    '***********************************************************************
    '!!! The following events will only be triggered if macro is enabled !!!
    '***********************************************************************
    If g_Macro.Ticking Then
     
        'Rare discovery messages
        If InStr(LCase(bstrMsg), "has discovered the") And (pColor = 5) Then
            MyDebug "clsACEvents.HandleConsoleMessage: OnRareFound: (" & pColor & ") " & Left(bstrMsg, Len(bstrMsg) - 1)
            'Play the rare notification sounds
            If g_ui.Options.chkAlertRare.Checked Then
                Call PlaySound(SOUND_RARE)
            End If
            'Call g_Macro.Loot.OnRareFound(bstrMsg)
        End If
    
        'The Corpse of Pestilence Rat is already in use by someone else
        If InStr(LCase(bstrMsg), "is already in use by someone else") And g_Macro.State = ST_LOOT Then
            Call g_Macro.Loot.IgnoreCurrentCorpse
        End If
        
        '==============================
        'Green Message (System Message)
        '==============================
        If pColor = 0 Or pColor = 1 Or pColor = 16 Or pColor = 18 Or pColor = 19 Or pColor = 23 Then
            'Fletching check
            If g_Macro.State = ST_FLETCHING Then
                If InStr(1, LCase(bstrMsg), "you make") Then
                    Call g_Macro.Fletcher.OnFletchingSuccessful
                ElseIf InStr(1, LCase(bstrMsg), "you fail to make") Then
                    Call g_Macro.Fletcher.OnFletchingFailed
                End If

            End If
            
            If InStr(LCase(bstrMsg), "missile attack hit the environment") Then
                Call OnHitEnvironment
            End If
            
            If g_ui.Options.chkAlertFellowDead.Checked And _
                InStr(LCase(bstrMsg), "your fellow") And InStr(LCase(bstrMsg), "has died") Then
                        Call PlaySound(SOUND_FELLOW_DEAD)
            End If
            
            If g_Macro.State = ST_FELLOW_CMD And InStr(LCase(bstrMsg), "already a memer of a fellowship") Then
                Call g_Macro.FellowCmd.notifyPlayer("You are already in a fellow, drop old fellow!")
                Call g_Macro.FellowCmd.ActionFinished
            End If
        End If
    
        '=====================================
        ' Vuln/Imperil/Yield Landing on Target
        '=====================================
        If (pColor = CHAT_SPELL_CASTING Or pColor = CHAT_SPELL_WORDS Or pColor = CHAT_SPELL_RESULTS) Then     'magic cast messages
            Msg = LCase(bstrMsg)
            
            'Look for "has expired" messages to make sure we don't drop any buffs
            '   The spell Cragstone's Will on Opal Blunt Baton has expired.
            '   Adja's Blessing has Expired
            If InStr(bstrMsg, "has expired") Then
                'See if it's an Item spell
                If InStr(bstrMsg, "The spell") Then
                    'Call g_Buffer.CheckExpiredItemSpell(Left(bstrMsg, Len(bstrMsg) - 1))
                Else
                    Call g_Buffer.CheckExpiredSpell(Left(bstrMsg, Len(bstrMsg) - 1)) ' get rid of char(10) on end
                End If
            End If
            
            If InStr(Msg, "prismatic taper") Then
                Call g_Macro.TapersCount(1)
                Call g_Spells.OnSpellCastComplete(True, "ACEvents: taper consumed")
            End If
            If InStr(Msg, "platinum scarab") Then
                Call g_Macro.PlatsCount(1)
            End If
            
            If InStr(bstrMsg, "you restore") Then
                Call g_Macro.SetFellowNeedsHealing(False)
                Call g_Spells.OnSpellCastComplete(True, "ACEvents: you restore")
                Call g_Macro.Combat.updateSecureTimer(4)
            End If
            
            If InStr(bstrMsg, "You say,") Then
                Call g_Spells.OnSpellCastBegin("ACEvents: spell words BEGIN: You say")
                'bFilterMsg = True
                If Valid(g_Macro.Combat.Target) And (g_Macro.State = ST_COMBAT_ATTACKING) Then
                    ' JSC - increment the MISSCOUNT
                    Dim missCount As Integer
                    missCount = g_Macro.Combat.Target.UserData(INT_MISSCOUNT) + 1
                    Call g_Macro.Combat.Target.SetUserData(INT_MISSCOUNT, missCount)
                    'MyDebug "clsACEvents.HandleConsoleMessage: +1 MISSCOUNT: " & missCount & " for: " & g_Macro.Combat.Target.Name
                End If
            ElseIf g_Macro.State = ST_REBUFF Then
                If InStr(bstrMsg, "You cast") Then
                    Call g_Buffer.CheckCastedSpell(bstrMsg)
                    Call g_Macro.Combat.updateSecureTimer(4)
                ElseIf InStr(Msg, "resists your spell") Then
                    Call g_Buffer.CheckUnenchantableItem(bstrMsg)
                    Call g_Macro.Combat.updateSecureTimer(4)
                End If
            ElseIf g_Macro.State = ST_BUDDYREBUFF Then
                If InStr(bstrMsg, "You cast") Then
                    Call g_BuddyBuffer.CheckCastedSpell(bstrMsg)
                    Call g_Macro.Combat.updateSecureTimer(4)
                ElseIf InStr(Msg, "resists your spell") Then
                    Call g_BuddyBuffer.CheckUnenchantableItem(bstrMsg)
                    Call g_Macro.Combat.updateSecureTimer(4)
                End If
            'ElseIf g_Macro.State = ST_FELLOWHEAL Then
            '    If InStr(bstrMsg, "you restore") Then
            '        Call g_Macro.SetFellowNeedsHealing(False)
            '        Call g_Spells.OnSpellCastComplete(True, "ACEvents: you restore")
            '    End If
            Else
            
                'Vuln landed
                If Valid(g_Spells.CurrentVuln) Then
                    If InStr(Msg, LCase("You cast " & Trim(g_Spells.CurrentVuln.SpellName))) Then
                        Call g_Macro.Combat.MageVulnLanded
                    End If
                End If
                
                'Spell got resisted ?
                If InStr(Msg, "resists your spell") Then
                    Call g_Macro.Combat.MageSpellGotResisted
                'Land Imperil
                ElseIf (InStr(Msg, "you cast gossamer flesh") Or InStr(Msg, "you cast imperil other")) Then
                    Call g_Macro.Combat.MageImperilLanded
                'Yield landing
                ElseIf (InStr(Msg, "you cast magic yield other") Or InStr(Msg, "you cast futility")) Then
                    Call g_Macro.Combat.MageYieldLanded
                ElseIf (InStr(Msg, "points with")) Then
                    'Our spell did damage to critter, reset the BLACKLIST count for this target
                    If Valid(g_Macro.Combat.Target) Then
                        Call g_Macro.Combat.Target.SetUserData(INT_MISSCOUNT, 0)
                        Call g_Macro.Combat.updateSecureTimer(4)
                        'MyDebug "clsACEvents.ConsoleMessage: reset MISSCOUNT : " & g_Macro.Combat.Target.Name
                        'MyDebug "clsACEvents.ConsoleMessage: Spell: " & Msg
                        Call g_Spells.OnSpellCastComplete(True, "ACEvents: Spell damage")
                    Else
                        MyDebug "clsACEvents.ConsoleMessage: points with, but no Valid Target"
                    End If
                End If

            End If
            
            If InStr(bstrMsg, "You cast") Then
                ' Must be Restamming or Healing
                Call g_Spells.OnSpellCastComplete(True, "ACEvents: You Cast")
            End If

        End If
        
        '=====================================
        ' Target Out of Range Message
        '=====================================
        If InStr(bstrMsg, "Target is out of range") Then
            MyDebug "Macro Target is out of range, reseting."
            Call g_Macro.Combat.StopCombat("HandleConsoleMessage - Target out of range")
        End If
    
        '=====================================
        ' Low or Out of Mana messages
        '=====================================
        If (pColor = CHAT_SPELL_CASTING Or pColor = CHAT_SPELL_WORDS) Then
            If (InStr(LCase(bstrMsg), "low on mana") Or InStr(LCase(bstrMsg), "out of mana")) Then
                Call g_Macro.setLowManaCheck
                bFilterMsg = False
            End If
        End If
        
        If (pColor = 0) And _
           (InStr(LCase(bstrMsg), "points of mana to") Or _
            InStr(LCase(bstrMsg), "no wielded items require mana") Or _
            InStr(LCase(bstrMsg), "item has no Mana to drain") Or _
            InStr(LCase(bstrMsg), "points of mana from the") Or _
            InStr(LCase(bstrMsg), "mana stone is destroyed")) Then
            Call g_Macro.doneManaCharge
        End If
    
        '=====================================
        'Patron/Vassel chat -- same color as spell casting... silly Turbine
        '=====================================
        If (pColor = CHAT_SPELL_CASTING Or pColor = CHAT_SPELL_WORDS Or pColor = CHAT_SPELL_RESULTS) Then
            If InStr(LCase(bstrMsg), "patron") Or InStr(LCase(bstrMsg), "vassals") Then
                bFilterMsg = False
            End If
        End If
        
        '=====================================
        ' Logging off... message, let's do it!
        '=====================================
        If (pColor = 0) And InStr(bstrMsg, "Logging off...") Then
            'Hmm, need to be sure to shutdown macro
            Call g_Macro.StopMacro
            Call g_Service.Logout("Found Logging off... message")
        End If
        
        '=====================================
        ' DOT tracking stuff
        '=====================================
        If g_ui.Macro.chkEnableDOT.Checked Then
            'MyDebug "clsACEvents.HandleConsoleMessage:(" & pColor & ") " & Left(bstrMsg, Len(bstrMsg) - 1)
            
            If InStr(bstrMsg, vbLf) Then
                Msg = Replace(bstrMsg, vbLf, "")
                'MyDebug "Detected vbLf in bstrMsg: " & bstrMsg
            End If
            
            'Find out what kinda damage it is
            If (pColor = CHAT_SPELL_CASTING) Or (pColor = CHAT_SPELL_RESULTS) Then
                'Critical hit! You bash Olthoi Ripper for 786 points with Shock Arc VII.
                'You bash Olthoi Ripper for 516 points with Shock Arc VII.
                    'Olthoi Ripper resists your spell
                '
                'White Phyntos Wasp chills you for 21 points with Frost Bolt IV.
                'White Phyntos Wasp numbs you for 11 points with Frost Volley IV.
                    'You resist the spell cast by White Phyntos Wasp
                'MyDebug "clsACEvents.HandleConsoleMessage:(" & pColor & ") " & Left(bstrMsg, Len(bstrMsg) - 1)

                If (InStr(Msg, "you for") And InStr(Msg, "points with")) Then
                    Call g_DOT.takeSpellDamage(Msg)
                ElseIf InStr(Msg, "points with") Or InStr(Msg, "resists your spell") Then
                    Call g_DOT.giveSpellDamage(Msg)
                End If
            
            'ElseIf (pColor = CHAT_ATTACKING_MONSTER) Then
            '    'You cut Ebon Mattekar for 67 points of slashing damage
            '        'Banderling Antagonist evaded your attack.
            '    If InStr(Msg, "point of") Or InStr(bstrMsg, "points of") Or InStr(Msg, "evaded your attack") Then
            '        MyDebug "Calling giveMeleeDamage"
            '        Call g_DOT.giveMeleeDamage(Msg)
            '    End If
            'ElseIf (pColor = CHAT_MONSTER_ATTACKING_US) Then
            '    'Critical hit! Olthoi Ripper nicks your lower leg for 10 points of piercing damage!
            '    'Olthoi Ripper nicks your upper leg for 4 points of piercing damage!
            '        'You evaded Olthoi Slasher!
            '    If (InStr(Msg, "your") And (InStr(Msg, "points of") Or InStr(Msg, "point of"))) Or InStr(Msg, "You evaded") Then
            '        Call g_DOT.takeMeleeDamage(Msg)
            '    End If
            End If
        End If
    
    Else
        ' Never filter if not running
        bFilterMsg = False
        
    End If  'macro active

    
    'Relay to IRC user/channel
    If Not bFilterMsg Then
    
        If Not InStr(bstrMsg, vbLf) And Not InStr(bstrMsg, vbCr) And Not InStr(bstrMsg, vbCrLf) Then
            ircString = bstrMsg & vbLf
        Else
            ircString = bstrMsg
        End If
        Call g_RemoteCmd.RemoteRedirectChatToIRC(ircString, pColor)
        
        'If Not (pColor = 27) And Not (pColor = 28) And Not (pColor = 29) And _
        '   Not (pColor = 7) And Not (pColor = 11) And Not (pColor = 17) And Not (pColor = 21) Then
        '    locDebug "clsACEvents.HandleConsoleMessage:(" & pColor & ") " & Left(bstrMsg, Len(bstrMsg) - 1)
        'End If
    Else
        locDebug "clsACEvents:Filtered: (" & pColor & ") " & bstrMsg
    End If

Fin:
    HandleConsoleMessage = bFilterMsg
    Exit Function
ErrorHandler:
    PrintErrorMessage "HandleConsoleMessage - " & Err.Description
    Resume Fin
End Function


'#####################################################################################
'#
'#                           Network Message Handler
'#
'#####################################################################################

'Public Sub HandleNetworkMessages(ByVal pMsg As DecalNet.IMessage2)
'On Error GoTo ErrorHandler
'
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "HandleNetworkMessages - " & Err.Description
'    Resume Fin
'End Sub

Private Sub m_tmrClockChecker_OnTimeout()

End Sub


'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsACEvents] " & DebugMsg, bSilent)
    End If
End Sub
