VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#####################################################################################
'#
'#                                  CONSTANTS
'#
'#####################################################################################

Private Const DEBUG_ME = False

Private Const DEFAULT_MACRO_TIMER_INTERVAL = 300   'msec
Private Const WAYPOINT_DELAY = 2                   'seconds
Private Const CAST_TIMER = 0.1                     'seconds

Private Const LOGOUT_CHECK_INTERVAL = 2     'check every X seconds
Private Const LOOT_CHECK_INTERVAL = 2       'check for loot every 2 seconds in idle mode
Private Const TARGET_SCAN_INTERVAL = 2      'check for targets every X seconds when in idle mode unless
                                            'timer gets reset when creatures are spawning nearby
Private Const SALVAGER_CHECK_INTERVAL = 5   'check if we need to autosalvage every 5 seconds

Private Const MAX_NAV_FAIL = 200            'maximum number of navigation/loot failures before reset.

Public Enum eMacroStates
    ST_IDLE = 0
    ST_COMBAT
    ST_REBUFF
    ST_BUDDYREBUFF
    ST_HEALING
    ST_RESTAM
    ST_FLETCHING
    ST_STICK_NAV
    ST_NAV
    ST_FELLOW_CMD
    ST_LOOT
    ST_SALVAGE
    ST_BUYSELL
    ST_MANACHARGE
End Enum

Public Enum eMacroSubStates
    SUBST_NONE = 0
    SUBST_CHANGE_COMBATMODE
    SUBST_EQUIP
End Enum

Public Enum eMacroModes
    MODE_NORMAL = 0
    MODE_REBUFF
    MODE_SURVIVE
End Enum

'Order important (for combobox listindex)
Public Enum eMacroCombatType
    TYPE_CASTER = 0
    TYPE_ARCHER
    TYPE_MELEE
    TYPE_NOTYPE
End Enum

Public Enum eRegenActions
    REGEN_REVITALIZE
    REGEN_HEALTH_TO_MANA
    REGEN_DRAIN_HEALTH
    REGEN_DRAIN_STAM
    REGEN_DRAIN_MANA
End Enum

Private Type TypeFiveMinCounter
  XP() As Double
  startXP() As Double
  Position As Long
  Count As Long
End Type

Private FiveMinXPCounter As TypeFiveMinCounter

'#####################################################################################
'#
'#                                DATA MEMBERS
'#
'#####################################################################################

Public PostCastDelayTimer As clsTimer
Private m_tmrBusyDelay As clsTimer
Private m_tmrNextLogoutCheck As clsTimer
Private WithEvents m_tmrWaypointDelay As clsTimer
Attribute m_tmrWaypointDelay.VB_VarHelpID = -1
Private WithEvents m_tmr5MinTimer As clsTimer
Attribute m_tmr5MinTimer.VB_VarHelpID = -1
Private WithEvents m_tmrCastSpell As clsTimer
Attribute m_tmrCastSpell.VB_VarHelpID = -1

Private m_iState As eMacroStates
Private m_iSubState As eMacroSubStates
Private m_iMode As Integer
Private m_iCombatType As eMacroCombatType
Private m_bRunning As Boolean
Private m_bPaused As Boolean
Private m_adminPaused As Boolean
Private m_fellowNeedsHealing As Boolean
Private m_fellowNeedsStam As Boolean
Public m_bOKToCast As Boolean ' Triggered by m_tmrCastSpell

'Timers / display / Stats
Private m_dStartTime As Double
Private m_dElapsedTime As Double
Private m_dLastResetTime As Double
Private m_dTimeSpent As Double      'Time since last reset

Private m_sTimeSpent As String      'formated string of the time since last reset
'Private m_sElapsedTime As String    'formated string of the elapsed time
Private m_sNextRebuff As String     'formated string of the time remaining until next rebuff

Private m_dXpAtStart As Double
Private m_dXpHour As Double
Private m_dXpEarned As Double
Private m_dTotalXp As Double
Private m_dFiveMinXPAvg As Double
Private m_dFiveMinXpHour As Double

Private m_sTimeUntilLevel As String 'formated string of the time remaining until next level
Private m_sXpHour As String         'formated string of the Xp/Hour
Private m_sTotalXp As String        'formated string of the current total Xp
Private m_sXpEarned As String       'formated string of the total amount of xp earned
Private m_sXpMinute As String       'formated string of the 5 min rate of Xp

Private m_dTapersCount As Double
Private m_dTapersHour As Double
Private m_sTapersHour As String

Private m_dPlatsCount As Double
Private m_dPlatsHour As Double
Private m_sPlatsHour As String

Private m_dArrowsCount As Double
Private m_dArrowsHour As Double
Private m_sArrowsHour As String

Private m_tmrNextStatsUpdate As clsTimer       'use to update stats every 1 seconds

'Regen
Private m_RegenManaAction As eRegenActions
Private m_RegenStamAction As eRegenActions
Private m_bRegenVitals As Boolean           'restaming / replenishing mana
Private m_tmrFellowHealTimer As clsTimer    'check fellow health every 30 seconds
Private m_tmrManaChargeTimer As clsTimer    'mana charge timer

'Misc
Private m_bNeedArrows As Boolean
Private m_bNeedManaCharge As Boolean
Private m_bUsingManaCharge As Boolean
Private m_bDied As Boolean
Private m_bForceRebuff As Boolean
Private m_bForceBuddyRebuff As Boolean
Private m_tmrNextTargetScan As clsTimer
Private m_fStartHeading As Single

'Pending Fellowship Actions
Private m_objFellowTarget As acObject
Private m_iFellowAction As eFellowshipAction
Private m_FellowCollection As Dictionary

'Looting
Private m_tmrNextLootCheck As clsTimer
Public m_colChargeItems As colObjects 'List of all the charge items.
Private m_NavFails As Integer

'Salvager
Private m_tmrNextSalvageTime As clsTimer
Private m_bForceSalvage As Boolean
Private m_iNumItemsPickedUp As Integer  'number of items picked up since last salvaging

'Macro Extansion Modules
Private WithEvents m_CombatMode As clsMacroChangeCM
Attribute m_CombatMode.VB_VarHelpID = -1
Private WithEvents m_Equip As clsMacroEquip
Attribute m_Equip.VB_VarHelpID = -1
Private WithEvents m_Healing As clsMacroHealing
Attribute m_Healing.VB_VarHelpID = -1
Private WithEvents m_Combat As clsMacroCombat
Attribute m_Combat.VB_VarHelpID = -1
Private WithEvents m_Fletcher As clsMacroFletcher
Attribute m_Fletcher.VB_VarHelpID = -1
Private WithEvents m_Nav As clsNav
Attribute m_Nav.VB_VarHelpID = -1
Private WithEvents m_Restam As clsMacroRestam
Attribute m_Restam.VB_VarHelpID = -1
Public WithEvents m_FellowCmd As clsFellowCmd
Attribute m_FellowCmd.VB_VarHelpID = -1
Private WithEvents m_Loot As clsLoot
Attribute m_Loot.VB_VarHelpID = -1
Private WithEvents m_Salvager As clsSalvager
Attribute m_Salvager.VB_VarHelpID = -1

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
'On Error GoTo ErrorHandler
 
    Set m_tmrBusyDelay = CreateTimer
    Set PostCastDelayTimer = CreateTimer
    Set m_tmrWaypointDelay = CreateTimer
    Set m_tmrNextLogoutCheck = CreateTimer
    Set m_tmrNextLootCheck = CreateTimer
    Set m_tmrNextStatsUpdate = CreateTimer
    Set m_tmrNextTargetScan = CreateTimer
    Set m_tmrNextSalvageTime = CreateTimer
    Set m_tmrFellowHealTimer = CreateTimer
    Set m_tmrManaChargeTimer = CreateTimer
    Set m_tmr5MinTimer = CreateTimer
    Set m_tmrCastSpell = CreateTimer
    
    Call m_tmrBusyDelay.Reset
    Call PostCastDelayTimer.Reset
    Call m_tmrWaypointDelay.Reset
    Call m_tmrNextLogoutCheck.Reset
    Call m_tmrNextLootCheck.Reset
    Call m_tmrNextSalvageTime.Reset
    Call m_tmrFellowHealTimer.Reset
    Call m_tmrManaChargeTimer.Reset
    Call m_tmr5MinTimer.Reset
    Call m_tmrCastSpell.Reset
    
    Call ResetTargetScanTimer
    
    Call m_tmrCastSpell.SetNextTime(CAST_TIMER)
    Call m_tmrNextSalvageTime.SetNextTime(SALVAGER_CHECK_INTERVAL)
    Call m_tmrNextLootCheck.SetNextTime(LOOT_CHECK_INTERVAL)
    
    g_Core.TimerInterval = DEFAULT_MACRO_TIMER_INTERVAL
    
    m_iMode = MODE_NORMAL
    m_iCombatType = TYPE_CASTER
    m_iState = ST_IDLE
    m_bRunning = False
    m_bPaused = False
    m_adminPaused = False
    m_fellowNeedsHealing = False
    m_fellowNeedsStam = False
    
    Set m_objFellowTarget = Nothing
    
    'Macro modules
    Set m_CombatMode = New clsMacroChangeCM
    Set m_Equip = New clsMacroEquip
    Set m_Healing = New clsMacroHealing
    Set m_Combat = New clsMacroCombat
    Set m_Fletcher = New clsMacroFletcher
    Set m_Nav = New clsNav
    Set g_Nav = m_Nav
    Set m_Restam = New clsMacroRestam
    Set m_FellowCmd = New clsFellowCmd
    Set m_Loot = New clsLoot
    Set m_Salvager = New clsSalvager
    
    m_RegenManaAction = REGEN_REVITALIZE
    m_RegenStamAction = REGEN_REVITALIZE
    m_bRegenVitals = False
    m_bNeedArrows = False
    m_bNeedManaCharge = False
    m_bUsingManaCharge = False
    m_bDied = False
    m_bForceRebuff = False
    m_bForceBuddyRebuff = False
    m_bForceSalvage = False
    m_iNumItemsPickedUp = 0
    m_fStartHeading = 0
    
    'Timer display
    Call SetElapsedTime(1)
    Call UpdateNextRebuffDisplay(1)
    Call UpdateTimeUntilLevelDisplay(1)
    
    Set m_colChargeItems = New colObjects
    
    'Stats
    m_dLastResetTime = 1
    m_dStartTime = 1
    m_dXpAtStart = 0
    m_dXpHour = 0
    m_dXpEarned = 0
    m_dTotalXp = 0
    m_dFiveMinXPAvg = 0
    m_dFiveMinXpHour = 0
    Call UpdateStatsStrings
    
    m_dTapersCount = 0
    m_dTapersHour = 0
    m_dPlatsCount = 0
    m_dPlatsHour = 0
    m_dArrowsCount = 0
    m_dArrowsHour = 0
    
    m_NavFails = 0
    
    FiveMinXPCounter.Count = 30
    FiveMinXPCounter.Position = 0
  
    ReDim FiveMinXPCounter.XP(FiveMinXPCounter.Count) As Double
    ReDim FiveMinXPCounter.startXP(FiveMinXPCounter.Count) As Double
    
    Set m_FellowCollection = New Dictionary

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "Error in clsMacro_Initialize()"
    MsgBox "ENGINE ERROR - clsMacro_Initialize - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub Class_Terminate()

    Set m_tmrCastSpell = Nothing
    Set m_tmrBusyDelay = Nothing
    Set PostCastDelayTimer = Nothing
    Set m_tmrWaypointDelay = Nothing
    Set m_tmrNextLogoutCheck = Nothing
    Set m_tmrNextLootCheck = Nothing
    Set m_tmrFellowHealTimer = Nothing
    Set m_tmrNextSalvageTime = Nothing
    Set m_tmr5MinTimer = Nothing
    Set m_tmrManaChargeTimer = Nothing
    Set m_tmrNextStatsUpdate = Nothing
    Set m_tmrNextTargetScan = Nothing
    
    Set m_CombatMode = Nothing
    Set m_Equip = Nothing
    Set m_Healing = Nothing
    Set m_Combat = Nothing
    Set m_Fletcher = Nothing
    Set m_Restam = Nothing
    Set m_Nav = Nothing
    Set g_Nav = Nothing
    Set m_FellowCmd = Nothing
    Set m_Loot = Nothing
    Set m_Salvager = Nothing
    
    Set m_objFellowTarget = Nothing
    
    Set m_FellowCollection = Nothing
    Set m_colChargeItems = Nothing
End Sub

Private Sub ResetFellowAction()
    Set m_objFellowTarget = Nothing
    m_iFellowAction = FA_NONE
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get State() As Integer
    State = m_iState
End Property

Public Property Get SubState() As Integer
    SubState = m_iSubState
End Property

Public Property Get Mode() As Integer
    Mode = m_iMode
End Property

Public Property Get RegenVitals() As Boolean
    RegenVitals = m_bRegenVitals
End Property

Public Property Get Ticking() As Boolean
    Ticking = g_Core.Ticking
End Property

Public Property Get Active() As Boolean
    Active = Ticking And (Not m_bPaused)
End Property

Public Property Get Running() As Boolean
    Running = (Active And m_bRunning)
End Property

Public Property Get Paused() As Boolean
    Paused = m_bPaused
End Property

Public Property Get CombatType() As Integer
    CombatType = m_iCombatType
End Property

Public Property Get adminPaused() As Boolean
    adminPaused = m_adminPaused
End Property

Public Property Let adminPaused(ByVal iVal As Boolean)
    m_adminPaused = iVal
End Property

Public Property Let CombatType(ByVal iVal As Integer)
    m_iCombatType = iVal
End Property

Public Property Get Combat() As clsMacroCombat
    Set Combat = m_Combat
End Property

Public Property Get Fletcher() As clsMacroFletcher
    Set Fletcher = m_Fletcher
End Property

Public Property Get BusyDelayTimer() As clsTimer
    Set BusyDelayTimer = m_tmrBusyDelay
End Property

Public Property Get FellowCmd() As clsFellowCmd
    Set FellowCmd = m_FellowCmd
End Property

Public Property Get ElapsedTime() As Double
    ElapsedTime = m_dElapsedTime
End Property

Public Property Get ElapsedTimeString() As String
    'ElapsedTimeString = m_sElapsedTime
    ElapsedTimeString = m_sTimeSpent
End Property

Public Property Get NextRebuffTimeString() As String
    NextRebuffTimeString = m_sNextRebuff
End Property

Public Property Get TimeUntilNextLevelString() As String
    TimeUntilNextLevelString = m_sTimeUntilLevel
End Property

Public Property Get TotalXpString() As String
    TotalXpString = m_sTotalXp
End Property

Public Property Get XpHourString() As String
    XpHourString = m_sXpHour
End Property

Public Property Get XpMinuteString() As String
    XpMinuteString = m_sXpMinute
End Property

Public Property Get XpEarnedString() As String
    XpEarnedString = m_sXpEarned
End Property

Public Property Get NeedHealing(Optional ByVal sSource As String = "") As Boolean
    NeedHealing = m_Healing.NeedHealing(sSource)
End Property

Public Property Get Loot() As clsLoot
    Set Loot = m_Loot
End Property

Public Property Get Salvager() As clsSalvager
    Set Salvager = m_Salvager
End Property

Public Property Get Died() As Boolean
    Died = m_bDied
End Property

Public Property Let Died(ByVal bVal As Boolean)
    m_bDied = bVal
End Property

Public Property Get ForceSalvage() As Boolean
    ForceSalvage = m_bForceSalvage
End Property

Public Property Let ForceSalvage(ByVal bVal As Boolean)
    m_bForceSalvage = bVal
    If bVal Then Call m_tmrNextSalvageTime.ExpireNow
End Property

'Salvager - Number of items picked up since last salvaging
Public Property Get NumItemsPickedUp() As Integer
    NumItemsPickedUp = m_iNumItemsPickedUp
End Property

Public Property Get TapersHourString() As String
    TapersHourString = m_sTapersHour
End Property

Public Sub TapersCount(ByVal iVal As Integer)
    m_dTapersCount = m_dTapersCount + iVal
End Sub

Public Property Get PlatsHourString() As String
    PlatsHourString = m_sPlatsHour
End Property

Public Sub PlatsCount(ByVal iVal As Integer)
    m_dPlatsCount = m_dPlatsCount + iVal
End Sub

Public Property Get ArrowsHourString() As String
    ArrowsHourString = m_sArrowsHour
End Property

Public Sub ArrowsCount(ByVal iVal As Integer)
    m_dArrowsCount = m_dArrowsCount + iVal
End Sub

Public Sub SetFellowNeedsHealing(ByVal aBool As Boolean)
    m_fellowNeedsHealing = aBool
End Sub

Public Sub SetFellowNeedsStam(ByVal aBool As Boolean)
    m_fellowNeedsStam = aBool
End Sub

'#####################################################################################
'#
'#                               NAVIGATION
'#
'#####################################################################################

Private Sub m_Nav_OnRouteComplete()
    locDebug "m_Nav_OnRouteComplete"
    
    If Ticking Then
        locDebug "m_Nav_OnRouteComplete - Route over"
        If (m_iState = ST_NAV) Then
            PrintMessage "End of Route reached - disabling navigation. If you want the macro to loop between all the route waypoints, set the route navigation type to LOOP."
            g_ui.Macro.chkEnableNav.Checked = False
            Call GoIdle
        ElseIf (m_iState = ST_STICK_NAV) Then
            Call GoIdle
        End If
    End If
End Sub

Private Sub m_Nav_OnTargetLost()
    locDebug "m_Nav_OnTargetLost"
    If m_iState = ST_NAV Then
        locDebug "m_Nav_OnTargetLost - Follow Target lost"
        Call GoIdle
    End If
End Sub

Private Sub m_Nav_OnWaypointReached(ByVal wp As clsNavWaypoint)
    locDebug "m_Nav_OnWaypointReached -- FIRED !"
    
    If m_iState = ST_NAV Then
        'locDebug "m_Nav_OnWaypointReached - Macro reached waypoint : " & wp.Description
        Call CheckStopNav
        Call m_tmrWaypointDelay.SetNextTime(WAYPOINT_DELAY)
        Call GoIdle
    ElseIf m_iState = ST_STICK_NAV Then
        locDebug "m_Nav_OnWaypointReached - ST_STICK_NAV -- wp: " & wp.Description
        Call CheckStopNav
        Call m_tmrWaypointDelay.SetNextTime(WAYPOINT_DELAY)
        Call GoIdle
    End If
End Sub



'#####################################################################################
'#
'#                               OBJECT EVENTS
'#
'#####################################################################################

'Update - State engine
Friend Sub OnMacroTick()
On Error GoTo ErrorHandler

    'increase elapsed time
     If (g_Core.Time - m_dStartTime > 0) Then
1        Call SetElapsedTime(g_Core.Time - m_dStartTime)
        'MyDebug "OnMacroTick: " & CDbl(g_Core.Time - m_dStartTime)
     End If
     
2    Call UpdateNextRebuffDisplay
    
    'perform state-independant task(s)
3    Call OnUpdate
    
    If Not m_bPaused Then
        'run the current state task(s)
4        Call RunState
    End If
    
    
Fin:
    Exit Sub
ErrorHandler:
    'PrintErrorMessage "clsMacro.OnMacroTick() - " & Err.Description & " - line : " & Erl
    Resume Fin
End Sub

Private Sub m_CombatMode_OnCombatModeChanged()
    If m_iSubState = SUBST_CHANGE_COMBATMODE Then
        locDebug "Macro : combat state changed."
        Call ResetSubState
    Else
        If m_Healing.NeedHealing Then
            locDebug "Macro : combat state changed."
            Call ResetSubState
        Else
            MyDebug "m_CombatMode_OnCombatModeChanged fired, but SubState != SUBST_CHANGE_COMBATMODE"
        End If
    End If
End Sub

Private Sub m_Equip_OnItemEquipped()
    If m_iSubState = SUBST_EQUIP Then
        locDebug "Macro : item equipped."
        Call ResetSubState
    Else
        PrintWarning "m_Equip_OnItemEquipped fired, but SubState != SUBST_EQUIP"
    End If
End Sub

Private Sub m_Healing_OnHealingComplete()
    If m_iState = ST_HEALING Then
        locDebug "Macro : Healing Complete."
        Call GoIdle
    Else
        PrintWarning "m_Healing_OnHealingComplete fired, but State != ST_HEALING"
    End If
End Sub

Private Sub m_Combat_OnStopCombat()
    If m_iState = ST_COMBAT Then
        locDebug "clsMacro - m_Combat_OnStopCombat : current combat over, moving back to Idle"
        Call GoIdle
    Else
        'PrintWarning "m_Combat_OnStopCombat fired, but State <> ST_COMBAT - Ignoring"
    End If
End Sub

Private Sub m_Fletcher_OnFletchingOver()
    If m_iState = ST_FLETCHING Then
        locDebug "clsMacro - m_Fletcher_OnFletchingOver : current combat over, moving back to Idle"
        Call GoIdle
    Else
        'PrintWarning "m_Fletcher_OnFletchingOver fired, but State <> ST_FLETCHING - Ignoring"
    End If
End Sub

Private Sub m_Restam_OnRestamComplete()
    If m_iState = ST_RESTAM Then
        locDebug "clsMacro - m_Restam_OnRestamComplete : restaming done, moving back to Idle"
        Call GoIdle
    Else
        'PrintWarning "m_Restam_OnRestamComplete fired, but State <> ST_RESTAM - Ignoring"
    End If
End Sub

Private Sub m_FellowCmd_OnActionComplete()
    If m_iState = ST_FELLOW_CMD Then
        locDebug "clsMacro - m_FellowCmd_OnActionComplete : action done, moving back to Idle"
        Call ResetFellowAction
        Call GoIdle
    Else
        'PrintWarning "m_FellowCmd_OnActionComplete fired, but State <> ST_FELLOW_CMD - Ignoring"
    End If
End Sub

Private Sub m_Loot_OnLootingComplete()
    If m_iState = ST_LOOT Then
        locDebug "clsMacro - m_Loot_OnLootingComplete : looting done, moving back to Idle"
        Call GoIdle
    Else
        'PrintWarning "m_Loot_OnLootingComplete fired, but State <> ST_LOOT - Ignoring"
    End If
End Sub

Private Sub m_Loot_OnItemLooted(objItem As DarksideFilter.acObject)
    m_iNumItemsPickedUp = m_iNumItemsPickedUp + 1
    locDebug "clsMacro.m_Loot_OnItemLooted -- m_iNumItemsPickedUp = " & m_iNumItemsPickedUp
End Sub

Private Sub m_Salvager_OnSalvageStopped(ByVal bComplete As Boolean)
    If m_iState = ST_SALVAGE Then
        locDebug "clsMacro - m_Salvager_OnSalvageComplete : salvaging done, moving back to Idle"
        m_bForceSalvage = False
        If bComplete Then m_iNumItemsPickedUp = 0 'reset num of items picked up
        Call GoIdle
    Else
        'PrintWarning "m_Salvager_OnSalvageComplete fired, but State <> ST_SALVAGE - Ignoring"
    End If
End Sub

'#####################################################################################
'#
'#                               PRIVATE
'#
'#####################################################################################

Private Sub OnUpdate()
On Error GoTo ErrorHandler

    Call g_ui.Main.UpdateStatus
    
    If m_tmrNextStatsUpdate.Expired Then
        Call UpdateStats
        Call g_ui.Main.UpdateStats
        Call m_tmrNextStatsUpdate.SetNextTime(5)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnUpdate() - " & Err.Description
    Resume Fin
End Sub

Public Sub GoIdle()
On Error GoTo ErrorHandler

    Call m_tmrNextLootCheck.Reset
    Call g_ui.Main.ResetBuffsProgressBar
    Call ResetTargetScanTimer
    Call SetState(ST_IDLE)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.GoIdle() - " & Err.Description
    Resume Fin
End Sub

Private Sub GoNav()
On Error GoTo ErrorHandler

    'PrintErrorMessage ("Nav")

    'fixme : force to loop?
    If g_Nav.NavType = NAVTYPE_NONE Then
        g_Nav.NavType = NAVTYPE_REVERSE
    End If
    
    If g_Nav.ResumeRoute(True) Then
        Call SetState(ST_NAV)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.GoNav() - " & Err.Description
    Resume Fin
End Sub

Private Function GoNavSticky() As Boolean
On Error GoTo ErrorHandler

    'Only return to sticky point if critters are around or there are no more corpses to loot
    If m_Loot.canLoot And (m_Combat.Target Is Nothing) Then
        If m_tmrNextTargetScan.Expired Then
            ' First check for Targets within Attack Range
            Dim objTarget As acObject
            If Not FindBestTarget(objTarget) Then
                GoNavSticky = False
                Exit Function
            End If
        Else
            GoNavSticky = False
            Exit Function
        End If
    End If

    If g_Nav.MoveToSticky() Then
        GoNavSticky = True
        Call SetState(ST_STICK_NAV)
    Else
        GoNavSticky = False
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsMacro.GoNavSticky() - " & Err.Description
    Resume Fin
End Function

'GoBuddyRebuff
Private Function GoBuddyRebuff() As Boolean
On Error GoTo ErrorHandler

    If Not ValidRangeTo(g_buffBuddy, g_ui.Macro.txtAssistRange.Text) Then
        PrintMessage "Buff Buddy out of range, disabling buff buddy!"
        Call g_BuddyBuffer.StopService
        Set g_buffBuddy = Nothing
        GoBuddyRebuff = False
    End If
    
    If ValidState(TYPE_CASTER) Then
        g_BuddyBuffer.mNumBuffsToCast = g_BuddyBuffer.BuffQueue.Count
        Call SetState(ST_BUDDYREBUFF)
        GoBuddyRebuff = True
    Else
        locDebug "clsMacro.GoBuddyRebuff: NOT ValidState(TYPE_CASTER)"
        GoBuddyRebuff = False
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsMacro.GoBuddyRebuff() - " & Err.Description
    Resume Fin
End Function


Private Sub GoRebuff()
On Error GoTo ErrorHandler

    If ValidState(TYPE_CASTER) Then
        g_Buffer.mNumBuffsToCast = g_Buffer.BuffQueue.Count
        g_ui.Main.progBuffs.MaxValue = g_Buffer.mNumBuffsToCast
        g_ui.Main.progBuffs.PostText = "/" & g_Buffer.mNumBuffsToCast
        g_ui.Main.progBuffs.Value = 0
        g_ui.Main.progBuffs.DecalDrawText = True
        Call SetState(ST_REBUFF)
    Else
        locDebug "clsMacro.GoRebuff: NOT ValidState(TYPE_CASTER)"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.GoRebuff() - " & Err.Description
    Resume Fin
End Sub

Private Function GoHealing() As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    bRet = m_Healing.StartHealing
    If bRet Then
        Call SetState(ST_HEALING)
        Call RunState
    End If
    
Fin:
    GoHealing = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.GoHealing - " & Err.Description
    Resume Fin
End Function

Private Function GoRestam() As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    bRet = m_Restam.StartRestam
    If bRet Then
        Call SetState(ST_RESTAM)
        Call RunState
    End If
    
Fin:
    GoRestam = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.GoRestam - " & Err.Description
    Resume Fin
End Function

Public Sub GoCombat(objTarget As acObject)
On Error GoTo ErrorHandler

    If Not Valid(objTarget) Then
        PrintErrorMessage "GoCombat : invalid objTarget"
        objTarget = Nothing
        Exit Sub
    End If
    
    If m_Combat.EngageCombat(objTarget, "GoCombat") Then
        Call SetState(ST_COMBAT)
        Call RunState
    Else
        MyDebug "GoCombat - EngageCombat failed"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.GoCombat() - " & Err.Description
    Resume Fin
End Sub


Private Function GoLooting() As Boolean
On Error GoTo ErrorHandler

    If m_Loot.StartLooting(True, g_ui.Loot.chkLootGround.Checked) Then
        GoLooting = True
        Call SetState(ST_LOOT)
        Call RunState
    Else
        GoLooting = False
        'MyDebug "clsMacro.GoLooting - StartLooting failed"
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsMacro.GoLooting() - " & Err.Description
    GoLooting = False
    Resume Fin
End Function

Private Function GoSalvage() As Boolean
On Error GoTo ErrorHandler

    If m_Salvager.StartSalvaging Then
        GoSalvage = True
        Call SetState(ST_SALVAGE)
        Call RunState
    Else
        GoSalvage = False
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsMacro.GoSalvage - " & Err.Description
    GoSalvage = False
    Resume Fin
End Function

Public Sub GoFletching()
On Error GoTo ErrorHandler

    If m_Fletcher.StartFletching(g_Data.Arrows, g_Data.ArrowHead, g_Data.ArrowShaft, "clsMacro.GoFletching") Then
        LogEvent "Fletching " & g_Data.Arrows & " with " & g_Data.ArrowHead & " and " & g_Data.ArrowShaft
        Call SetState(ST_FLETCHING)
    Else
        MyDebug "GoFletching - StartFletching failed"
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.GoFletching() - " & Err.Description
    Resume Fin
End Sub

Private Sub OnRebuffComplete()
On Error GoTo ErrorHandler

    'if in rebuff only mode, stop the macro
    If m_iMode = MODE_REBUFF Then
        PrintMessage "Buff Cycle Completed."
        Call StopMacro
    Else
        locDebug "OnRebuffComplete - Buff cycle completed, moving back to idle mode"
        m_bForceRebuff = False
        Call GoIdle
        Call RunState
    End If
    
    Call g_ui.Main.ResetBuffsProgressBar
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnRebuffComplete() - " & Err.Description
    Resume Fin
End Sub

Public Sub RecruitPlayer(objPlayer As acObject)
On Error GoTo ErrorHandler
    
    If Not Valid(objPlayer) Then
        PrintMessage "Fellowship recruit request, but target was to far away"
        Exit Sub
    End If
    
    If g_Objects.Fellowship.Active Then
        If m_iState <> ST_FELLOW_CMD Then
            SendTell objPlayer.Name, "Your fellowship request has been received. Please wait around me until Im ready to recruit you."
            Set m_objFellowTarget = objPlayer
            m_iFellowAction = FA_RECRUIT
        End If
    End If
        
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.RecruitPlayer() - " & Err.Description
    Resume Fin
End Sub

Public Function GoFellowshipCmd() As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    Select Case m_iFellowAction
        Case FA_RECRUIT
            If m_FellowCmd.RecruitPlayer(m_objFellowTarget) Then
                Call SetState(ST_FELLOW_CMD)
                bRet = True
            Else
                Call ResetFellowAction
                bRet = False
            End If
            
        Case Else
            'not supported, reset
            Set m_objFellowTarget = Nothing
            m_iFellowAction = FA_NONE
            bRet = False
    
    End Select

Fin:
    GoFellowshipCmd = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.GoFellowshipCmd() - " & Err.Description
    Resume Fin
End Function

' See if any fellow members nearby need healing
Private Function GoFellowHeal(Optional ByVal bStam As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
   
    bRet = False

    If Not g_Objects.Fellowship.Active Then GoTo Fin

    Dim objFellow As acObject
    
    Set objFellow = Nothing
    
    For Each objFellow In g_Objects.Fellowship
        
        If Not Valid(objFellow) Then GoTo NextObj
        If (objFellow.Name = g_Filters.playerName) Then GoTo NextObj
        
        'MyDebug "clsMacro.GoFellowHeal: " & objFellow.Name & ":health:" & objFellow.Health & " out of " & objFellow.MaxHealth
        
        If (objFellow.Health <= 0) Then GoTo Fin
        
        If (bStam) Then
            If (objFellow.Stamina <= GetPercent(objFellow.MaxStamina, g_Data.FellowStamThreshold)) Then
                If Not ValidRangeTo(objFellow, g_ui.Macro.txtAssistRange.Text) Then
                    'MyDebug "clsMacro.GoFellowHeal: out of range: " & objFellow.Name & " : " & objFellow.GetRange
                    GoTo NextObj
                End If
                'Heal them!
                MyDebug "clsMacro.GoFellowHeal: Staming " & objFellow.Name & " :guid: " & objFellow.Guid
                Call g_Spells.Cast_StamOther(objFellow.Guid)
                m_fellowNeedsStam = False
                bRet = True
            End If
        Else
            If (objFellow.Health <= GetPercent(objFellow.MaxHealth, g_Data.FellowHealthThreshold)) Then
                'MyDebug "clsMacro.GoFellowHeal: " & objFellow.Name & ":health:" & objFellow.Health & " at " & GetPercent(objFellow.MaxHealth, g_Data.FellowHealthThreshold)
            
                If Not ValidRangeTo(objFellow, g_ui.Macro.txtAssistRange.Text) Then
                    'MyDebug "clsMacro.GoFellowHeal: out of range: " & objFellow.Name & " : " & objFellow.GetRange
                    GoTo NextObj
                End If
                'Heal them!
                MyDebug "clsMacro.GoFellowHeal: healing " & objFellow.Name & " :guid: " & objFellow.Guid
                Call g_Spells.Cast_HealOther(objFellow.Guid)
                m_fellowNeedsHealing = False
                bRet = True
            End If
        End If
        
NextObj:
    Next objFellow
    'Next i
    
    Set objFellow = Nothing

Fin:
    GoFellowHeal = bRet
    Exit Function
ErrorHandler:
    GoFellowHeal = False
    PrintErrorMessage "clsMacro.GoFellowHeal() - " & Err.Description
    Exit Function
End Function


'***************************************
' SelectRegenMethod
'
' Choose the stamina & mana regeneration
' method (revit or drain)
'***************************************
Private Sub SelectRegenMethod()

    Select Case m_iState
            
        Case Else
            m_RegenManaAction = REGEN_REVITALIZE
            m_RegenStamAction = REGEN_REVITALIZE
            
    End Select
    
End Sub


'***************************************
' RechargingManaStam
'
' In magic mode, check if the macro needs to regen
' its mana/stamina. Returns true if it does, and
' casts the appropriate spell (revit/stam2mana, etc)
'***************************************
Private Function RechargingManaStam(Optional bMeleeMode As Boolean = False) As Boolean
On Error GoTo ErrorHandler

    'Choose what method to use to regen vitals
    Call SelectRegenMethod

    'Cast the appropriate regen spell if needed
    m_bRegenVitals = Vitals.CastManaStamRegen(False, m_RegenManaAction, m_RegenStamAction)

Fin:
    RechargingManaStam = m_bRegenVitals
    Exit Function
ErrorHandler:
    m_bRegenVitals = False
    PrintErrorMessage "clsMacro.RechargingManaStam - " & Err.Description
    Resume Fin
End Function


'***************************************
' isReadyToCast
'
' Returns true if we're not busy or/and already
' casting a spell
'***************************************
Public Function isReadyToCast(Optional ByVal bTurbo As Boolean = False) As Boolean
Dim bRet As Boolean
    
    'locDebug "clsMacro.isReadyToCast bTurbo: " & bTurbo & " IsBusy: " & IsBusy(bTurbo) & "  g_spells.Casting: " & g_Spells.Casting
    
    If IsBusy(bTurbo) Or g_Spells.Casting Then
        locDebug "clsMacro.isReadyToCast: IsBusy: " & IsBusy(bTurbo) & " : or g_Spells.Casting"
        bRet = False
    ElseIf Not (g_Spells.tmrPrecastDelay.Expired) Then
        locDebug "clsMacro.isReadyToCast: tmrPrecastDelay not Expired"
        bRet = False
    'ElseIf (g_Spells.LastCastAttempt + 1) > g_Core.Time Then
    '    locDebug "clsMacro.isReadyToCast: LastCastAttemp is less that g_Core.Time"
    '    'Limit to one try per second
    '    bRet = False
    'ElseIf Not (g_Macro.PostCastDelayTimer.Expired) Then
    '    locDebug "clsMacro.isReadyToCast: tmrPostCastDelayTimer not Expired"
    '    bRet = False
    Else
        'locDebug "clsMacro.isReadyToCast: Not RechargingManaStam"
        bRet = Not RechargingManaStam
    End If
    
    'locDebug "clsMacro.isReadyToCast: " & bRet
    
    isReadyToCast = bRet
End Function

'***************************************
' CheckStopNav
'
' Stop running if we are (navigation)
'***************************************
Public Sub CheckStopNav()
    If g_Nav.Running Then
        locDebug "CheckStopNav : stopping navigaiton"
        Call g_Nav.NavStop
    End If
End Sub


'***************************************
' RunState
'
' Main macro function - Runs the current state
'***************************************
Private Sub RunState()
On Error GoTo ErrorHandler

    'If m_PluginEnabled = False Then GoTo Fin
    
    Dim objTarget As acObject
    
    'don't do anything if in portal space
    If g_ds.InPortalSpace Then GoTo Fin
    
    'is the macro running?
    If Not m_bRunning Then GoTo Fin
    
    'is the macro paused?
    If m_bPaused Then GoTo Fin
    
    'Check if we need to logout
    If m_tmrNextLogoutCheck.Expired Then
        'If (g_AntiBan.m_bLogOut) Then
        '    Call g_Hooks.CastSpell(2023, 0)
        'End If
    
        Dim sReason As String
        If MustLogOut(sReason) Then
            MyDebug "RunState - MustLogOut returned true !"
            Call StopMacro
            Call g_Service.Logout(sReason)
            GoTo Fin
        End If
        Call m_tmrNextLogoutCheck.SetNextTime(LOGOUT_CHECK_INTERVAL)
    End If
    
    'If we're busy, don't do anything
    If IsBusy(TurboMode) Then
        locDebug "clsMacro.RunState: IsBusy(TurboMode) is true!"
        GoTo Fin
    End If
    
    'make sure we're not moving if we are not in ST_NAV state
    If (m_iState <> ST_NAV) And (m_iState <> ST_STICK_NAV) Then Call CheckStopNav
    
    ' Check for queued spells
    If g_Spells.c_SpellQueue.Count > 0 And m_bOKToCast Then
        locDebug "Queue has " & g_Spells.c_SpellQueue.Count & " spells"
        'Call g_Spells.c_SpellQueue.Display
        
        Dim o_Spell As clsSpellQueueItemSimple
        
        Set o_Spell = g_Spells.c_SpellQueue.ReadTop
                
        'PrintErrorMessage "Casting..."
        Call g_Hooks.CastSpell(o_Spell.SpellID, o_Spell.TargetGUID)
        
        Call g_Spells.c_SpellQueue.Pop ' Pop the spell off the queue
        
        ' Reset casting state
        m_bOKToCast = False
        m_tmrCastSpell.SetNextTime (CAST_TIMER)
        
        GoTo Fin
    End If
    
    'Handle sub states first
    If m_iSubState <> SUBST_NONE Then
        Select Case m_iSubState
                
            Case SUBST_CHANGE_COMBATMODE
                Call m_CombatMode.RunState
                    
            Case SUBST_EQUIP
                Call m_Equip.RunState
                    
            Case Else
                PrintWarning "RunState : Unhandled SubState " & m_iSubState
                Call ResetSubState
                    
            End Select
    Else
        Select Case m_iState
            Case ST_IDLE
                If ACBusy Then GoTo Fin
                
                locDebug "ST_IDLE: Idle"
                                
                'check to see if we need to loot a rare
                If g_bLootRare Then
                    locDebug "ST_IDLE : Looting Rare..."
                    If GoLooting Then
                        locDebug "ST_IDLE : Starting Looting..."
                        GoTo Fin
                    End If
                End If
                
                'check if we need healing
                If m_Healing.NeedHealing("ST_IDLE") Then
                    If GoHealing Then GoTo Fin
                End If
                
                'check to see if fellow members nearby need healing
                If g_ui.Macro.chkHealFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsHealing Then
                    locDebug "ST_IDLE: checking for fellow members in need of healing"
                    If GoFellowHeal Then GoTo Fin
                End If
                
                'check to see if fellow members nearby need Staming
                If g_ui.Macro.chkStamFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsStam Then
                    locDebug "ST_IDLE: checking for fellow members in need of Staming"
                    If GoFellowHeal(True) Then GoTo Fin
                End If
                
                'check to see if we need to loot a rare
                If g_bLootRare Then
                    locDebug "ST_IDLE : Looting Rare..."
                    If GoLooting Then
                        locDebug "ST_IDLE : Starting Looting..."
                        GoTo Fin
                    End If
                End If
                
                'check if we need to recruit someone
                If (m_iFellowAction = FA_RECRUIT) And Valid(m_objFellowTarget) Then
                    If GoFellowshipCmd Then GoTo Fin
                End If
                
                'check if we need to move back to sticky spot
                If g_ui.Macro.chkEnableSticky.Checked And g_Nav.isStickySet And g_Nav.stickyRunTimeout Then
                    locDebug "ST_IDLE: moving to Sticky Point"
                    If GoNavSticky Then GoTo Fin
                End If
                
                'check if we need to buff
                If (g_ui.Buffs.chkEnableBuffing.Checked Or m_bForceRebuff) And (g_Buffer.BuffQueue.Count > 0) And (m_bOKToCast = True) Then
                    If ValidState(TYPE_CASTER) Then
                        Call GoRebuff
                    End If
                    GoTo Fin
                End If
                
                'check if we need to buff our little buddy
                If Valid(g_buffBuddy) And (g_ui.Buffs.chkEnableBuffBuddy.Checked Or m_bForceBuddyRebuff) And (g_BuddyBuffer.BuffQueue.Count > 0) And (m_bOKToCast = True) Then
                    If GoBuddyRebuff Then GoTo Fin
                End If
                
                'check if we need to reload arrows
                If (m_iCombatType = TYPE_ARCHER) And m_bNeedArrows Then
                    Call ReloadArrows
                    GoTo Fin
                End If
            
                'Need stam?
                'Check to make sure that restam method has not changed.
                Call m_Restam.ChooseBestRestamMethod 'Get new Restam Method.
                
                If (m_Restam.RestamMethod <> NO_RESTAM) And Vitals.NeedStamina(True) Then
                    If Not IsBusy Then
                        Call GoRestam
                    Else
                        MyDebug "ST_IDLE - Macro need stam but currently busy - waiting still..."
                    End If
                    GoTo Fin
                End If
                
                'Check if we need mana.
                If IsCaster Then
                    'MyDebug "Checking if we need more stamina..."
                    Call Vitals.CastManaStamRegen(False, m_RegenManaAction, m_RegenStamAction)
                End If
                
                ' See if we need to change our Items with a Mana Charge
                If m_bNeedManaCharge And g_ui.Macro.chkUseManaCharge.Checked Then
                    Call goManaCharge
                    GoTo Fin
                End If
                
                
                If IsCaster And Not (g_ui.Loot.chkLootForcePeaceMode.Checked) Then
                    If Not isReadyToCast(TurboMode) Then 'make sure we keep our mana/stam up
                        locDebug "clsMacro.RunState: ST_IDLE: Not isReadyToCast()"
                        GoTo Fin
                    End If
                End If
                
                'Check if loot priority is boosted.
                If g_ui.Loot.chkBoostLootPriority.Checked _
                And g_ui.Loot.chkEnableLooting.Checked _
                And (Not g_Objects.Items.BackpackFull) _
                And m_tmrNextLootCheck.Expired Then
                    'locDebug "ST_IDLE : Checking Loot..."
                    If GoLooting Then
                        locDebug "ST_IDLE : Starting Looting..."
                        GoTo Fin
                    End If
                    Call m_tmrNextLootCheck.SetNextTime(LOOT_CHECK_INTERVAL)
                End If
                
                'PrintErrorMessage ("Debug: Idle 3")
                    
                'Look up for targets
                If m_Combat.Target Is Nothing Then
                    'SPK - Added a scan timer so that we don't do more CPU-expansive scans than
                    'what's necessary
                    If m_tmrNextTargetScan.Expired Then
                        Call m_tmrNextTargetScan.SetNextTime(TARGET_SCAN_INTERVAL)
                        ' First check for Targets within Attack Range
                        If FindBestTarget(objTarget) Then
                            Call GoCombat(objTarget)
                            GoTo Fin
                        ' Now check for non-vulned Targets within VulnRange
                        ElseIf g_ui.Macro.chkVuln.Checked Then
                            If FindNonDebuffedTarget(objTarget, g_ui.Macro.txtVulnRange.Text, "FindBestTarget") Then
                                Call GoCombat(objTarget)
                                GoTo Fin
                            End If
                        End If
                    End If
                        
                    '
                    'Not Target Found
                    '
                        
                    'Check if we should run the auto-salvager
                    If g_ui.Loot.chkEnableSalvage.Checked _
                    And m_tmrNextSalvageTime.Expired Then
                        If m_bForceSalvage _
                        Or g_Objects.Items.BackpackFull _
                        Or (g_ui.Loot.chkSalvageFrequency.Checked And (m_iNumItemsPickedUp >= g_Data.SalvagerFrequency)) Then
                            'locDebug "ST_IDLE : Salvaging"
                            Call m_tmrNextSalvageTime.SetNextTime(SALVAGER_CHECK_INTERVAL)
                            If GoSalvage Then
                                locDebug "ST_IDLE : Starting AutoSalvager..."
                                GoTo Fin
                            End If
                        End If
                        Call m_tmrNextSalvageTime.SetNextTime(SALVAGER_CHECK_INTERVAL)
                    End If
                        
                        
                    'Check if there's something to loot
                    If g_ui.Loot.chkEnableLooting.Checked _
                    And (Not g_Objects.Items.BackpackFull) _
                    And m_tmrNextLootCheck.Expired Then
                        'locDebug "ST_IDLE : Checking Loot..."
                        Call m_tmrNextLootCheck.SetNextTime(LOOT_CHECK_INTERVAL)
                        If GoLooting Then
                            locDebug "ST_IDLE : Starting Looting..."
                            GoTo Fin
                        End If
                    End If
                        
                    'Check if we can move to next waypoint. Don't interrupt looting for this.
                    'PrintErrorMessage ("Checking waypoint...")
                    
                    If m_iState = ST_LOOT Or m_Loot.getNavHold > 0 Then
                        
                        m_NavFails = m_NavFails + 1
                        
                        If m_NavFails >= MAX_NAV_FAIL Then ' Don't get stuck.
                            'Call PrintErrorMessage("Resetting nav timer.")
                            m_NavFails = 0
                            Call m_Loot.resetNavHold
                            
                            If g_ui.Macro.chkEnableNav.Checked _
                            And (g_Nav.Route.NumWP >= 0) _
                            And m_tmrWaypointDelay.Expired Then
                                locDebug "ST_IDLE : resuming route"
                                Call GoNav
                                GoTo Fin
                            End If
                        End If
                        
                        'PrintErrorMessage ("Looting not done.")
                        
                        'If (m_iState = ST_LOOT) Then
                        '    PrintErrorMessage ("State.")
                        'Else
                        '    PrintErrorMessage ("navHold")
                        'End If
                    Else
                        m_NavFails = 0
                        Call m_Loot.resetNavHold
                        
                        'PrintErrorMessage ("Looting done.")
                        
                        If g_ui.Macro.chkEnableNav.Checked _
                        And (g_Nav.Route.NumWP > 0) _
                        And m_tmrWaypointDelay.Expired Then
                            locDebug "ST_IDLE : resuming route"
                            Call GoNav
                            GoTo Fin
                        End If
                    End If
                Else
                    Call GoCombat(m_Combat.Target)  'resume combat
                End If
                    
                ' Make sure we are in combat state when Idle
                If Not (g_ui.Loot.chkLootForcePeaceMode.Checked) And (g_Hooks.CombatMode <> COMBATSTATE_PEACE) Then
                    Call ValidState
                End If
                
            Case ST_STICK_NAV   ' moving back to sticky point
                'check if we need healing
                If m_Healing.NeedHealing("ST_NAV") Then
                    If GoHealing Then GoTo Fin
                End If
                
                'Need stam ?
                Call m_Restam.ChooseBestRestamMethod 'Get new Restam Method.
                
                If Vitals.NeedStamina(IsMelee) And (m_Restam.RestamMethod <> NO_RESTAM) Then
                    Call GoRestam
                    GoTo Fin
                End If
                
                ' Check to see if we've been running to long
                If Not g_Nav.routeCheckerRunning Then
                    locDebug "ST_STICK_NAV: routeCheckerRunning false, GoIdle"
                    Call GoIdle
                End If
 
            Case ST_NAV 'a "in-motion" variation of ST_IDLE
                'check if we need healing
                If m_Healing.NeedHealing("ST_NAV") Then
                    If GoHealing Then GoTo Fin
                End If
                
                'Need stam ?
                Call m_Restam.ChooseBestRestamMethod 'Get new Restam Method.
                
                If Vitals.NeedStamina(IsMelee) And (m_Restam.RestamMethod <> NO_RESTAM) Then
                    Call GoRestam
                    GoTo Fin
                End If
                
                'Look for targets
                If m_Combat.Target Is Nothing Then
                    If FindBestTarget(objTarget) Then
                        locDebug "ST_NAV : found target while moving, moving to regular IDLE state for combat initialization"
                        'go back to regular IDLE mode to ensure we're in the proper equipment/vital state
                        Call GoIdle
                    End If
                Else
                    Call GoIdle
                End If
                
                ' Check to see if we've been running to long
                If Not g_Nav.routeCheckerRunning Then
                    locDebug "ST_NAV: routeCheckerRunning false, GoIdle"
                    Call GoIdle
                End If
                
            Case ST_COMBAT
                'check if we need healing
                If m_Healing.NeedHealing("ST_COMBAT") Then
                    If GoHealing Then GoTo Fin
                End If
                
                'Need stam (for melee/archer only) ?
                Call m_Restam.ChooseBestRestamMethod 'Get new Restam Method.
                
                If Vitals.NeedStamina(IsMelee) And (m_Restam.RestamMethod <> NO_RESTAM) Then
                    Call GoRestam
                    GoTo Fin
                End If
                
                'check to see if fellow members nearby need healing
                If g_ui.Macro.chkHealFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsHealing Then
                    locDebug "ST_IDLE: checking for fellow members in need of healing"
                    If GoFellowHeal Then GoTo Fin
                End If
                
                'check to see if fellow members nearby need Staming
                If g_ui.Macro.chkStamFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsStam Then
                    locDebug "ST_IDLE: checking for fellow members in need of Staming"
                    If GoFellowHeal(True) Then GoTo Fin
                End If
                
                If ValidState Then
                    Call m_Combat.RunState
                End If
                
                'Check if loot priority is boosted.
                If g_ui.Loot.chkBoostLootPriority.Checked _
                And g_ui.Loot.chkEnableLooting.Checked _
                And (Not g_Objects.Items.BackpackFull) _
                And m_tmrNextLootCheck.Expired Then
                    'locDebug "ST_IDLE : Checking Loot..."
                    If GoLooting Then
                        locDebug "ST_IDLE : Starting Looting..."
                        GoTo Fin
                    End If
                    Call m_tmrNextLootCheck.SetNextTime(LOOT_CHECK_INTERVAL)
                End If
                
            Case ST_RESTAM
                'check if we need healing
                If m_Healing.NeedHealing("ST_RESTAM") Then
                    If GoHealing Then GoTo Fin
                End If
                
                'Run restam state
                Call m_Restam.RunState
                
            Case ST_BUDDYREBUFF
                'check if we still need to be in rebuff mode
                If g_BuddyBuffer.BuffQueue.Count <= 0 Then
                    locDebug "clsMacro.RunState:ST_BUDDYREBUFF: no more buffs on the stack, moving back to IDLE"
                    Call OnRebuffComplete
                    GoTo Fin
                Else
                    'check if we need healing
                    If m_Healing.NeedHealing("ST_BUDDYREBUFF") Then
                        If GoHealing Then GoTo Fin
                    End If
                    
                    If (isReadyToCast(TurboMode)) And (m_bOKToCast = True) Then
                        locDebug "clsMacro.RunState:ST_BUDDYREBUFF: isReadyToCast true, CastNextSpell"
                        Call g_BuddyBuffer.CastNextSpell
                    Else
                        locDebug "clsMacro.RunState:ST_BUDDYREBUFF: isReadyToCast NOT true"
                    End If
                End If
                
            Case ST_REBUFF
                'check if we still need to be in rebuff mode
                If g_Buffer.BuffQueue.Count <= 0 Then
                    locDebug "clsMacro.RunState:ST_REBUFF: no more buffs on the stack, moving back to IDLE"
                    Call OnRebuffComplete
                    GoTo Fin
                Else
                    'check if we need healing
                    If m_Healing.NeedHealing("ST_REBUFF") Then
                        If GoHealing Then GoTo Fin
                    End If
                    
                    If (isReadyToCast(TurboMode)) And (m_bOKToCast = True) Then
                        locDebug "clsMacro.RunState:ST_REBUFF: isReadyToCast true, CastNextSpell"
                        Call g_Buffer.CastNextSpell
                    Else
                        locDebug "clsMacro.RunState:ST_REBUFF: isReadyToCast NOT true"
                    End If
                End If
            
            Case ST_HEALING
                Call m_Healing.RunState
                
            Case ST_FLETCHING
                If Not ACBusy Then
                    Call m_Fletcher.RunState
                End If
                
            Case ST_FELLOW_CMD
                If Not IsBusy Then
                    Call m_FellowCmd.RunState
                End If
                
            Case ST_LOOT
                'check to see if fellow members nearby need healing
                If g_ui.Macro.chkHealFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsHealing Then
                    locDebug "ST_IDLE: checking for fellow members in need of healing"
                    If GoFellowHeal Then GoTo Fin
                End If
                
                'check to see if fellow members nearby need Staming
                If g_ui.Macro.chkStamFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsStam Then
                    locDebug "ST_IDLE: checking for fellow members in need of Staming"
                    If GoFellowHeal(True) Then GoTo Fin
                End If
                
                If Not ACBusy Then
                    If g_ui.Loot.chkLootForcePeaceMode.Checked And (g_Hooks.CombatMode <> COMBATSTATE_PEACE) Then
                        Call g_Macro.RequestCombatStateChange(COMBATSTATE_PEACE)
                    End If
                    Call m_Loot.RunState
                End If
                
            Case ST_SALVAGE
                'check to see if fellow members nearby need healing
                If g_ui.Macro.chkHealFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsHealing Then
                    locDebug "ST_IDLE: checking for fellow members in need of healing"
                    If GoFellowHeal Then GoTo Fin
                End If
                
                'check to see if fellow members nearby need Staming
                If g_ui.Macro.chkStamFellow.Checked And g_Objects.Fellowship.Active And m_fellowNeedsStam Then
                    locDebug "ST_IDLE: checking for fellow members in need of Staming"
                    If GoFellowHeal(True) Then GoTo Fin
                End If
                
                If Not ACBusy Then
                    Call m_Salvager.RunState
                End If
            
            Case ST_BUYSELL
                If Not ACBusy Then
                    Call PhatLoot.RunState
                End If
            
            Case ST_MANACHARGE
                'check if we need healing
                If m_Healing.NeedHealing("ST_MANACHARGE") Then
                    If GoHealing Then GoTo Fin
                End If
                If Not ACBusy Then
                    Call UseManaCharge
                End If
                
        End Select
    End If
    
Fin:
    Set objTarget = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.RunState - " & Err.Description
    Resume Fin
End Sub

Private Function CheckMacroItem(ByVal sDescription As String, objItem As acObject, objEquipItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    bRet = True
    
    If Not Valid(objItem) Then
        'check to see if we have a wand in our hands
        If Valid(objEquipItem) Then
            PrintMessage "No default " & sDescription & " available : setting your current " & sDescription & " (" & objEquipItem.Name & ") as default. You can select another one from the Macro tab."
            
            Select Case sDescription
                Case ID_WAND
                    g_Data.WandGUID = objEquipItem.Guid
                
                Case ID_WEAPON
                    g_Data.WeaponGUID = objEquipItem.Guid
                    
                Case ID_SHIELD
                    g_Data.ShieldGUID = objEquipItem.Guid
                    
                Case ID_BOW
                    g_Data.BowGUID = objEquipItem.Guid
                    
            End Select
            Call g_ui.Macro.UpdateMacroItemsLabels
            
        Else
            PrintMessage "Please equip a " & sDescription & ", or set one in the Macro tab."
            bRet = False
        End If
    ElseIf Not g_Objects.Items.InInventory(objItem.Guid) Then
        PrintMessage "Please equip a valid " & sDescription & " from your inventory, or set one in the Macro tab."
        bRet = False
    End If
    
Fin:
    CheckMacroItem = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.CheckMacroItem - " & Err.Description
    Resume Fin
End Function

Private Sub OnMacroStart()
On Error GoTo ErrorHandler

1    'Init Strings
    m_sXpHour = "0"
    m_sTimeUntilLevel = "0"
    m_sXpHour = "0"
    m_sTotalXp = "0"
    m_sXpEarned = "0"
    m_sXpMinute = "0"
    m_sTapersHour = "0"
    m_sPlatsHour = "0"
    m_sArrowsHour = "0"
    m_sTimeSpent = "1"
    m_sNextRebuff = "1"

2   'turn on the macro timer
    m_bRunning = True
    m_bDied = False
    Call m_tmrNextStatsUpdate.Reset
    Call m_tmrNextStatsUpdate.SetNextTime(5)

    'reset counters
3    Call UpdateNextRebuffDisplay
4    m_dStartTime = g_Core.Time
5    m_dLastResetTime = 1
6    Call SetElapsedTime(1)
    
7    m_dTotalXp = g_ds.XpTracker.TotalXp
8    m_dXpAtStart = m_dTotalXp
9    Call UpdateStats
    
    'Update screen dimension
10    Call g_Window.UpdateDimensions
    
    'Start fellowship checker
11    Call m_tmrFellowHealTimer.SetNextTime(15)
    
    'Start 5 Min XP timer
12    Call m_tmr5MinTimer.SetNextTime(10)

13    Call Vitals.idManaStones
    
14    Call RunState
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnMacroStart - " & Err.Description & " line: " & Erl
    Resume Fin
End Sub

Public Sub ForceBuddyBuff()
On Error GoTo ErrorHandler

    Call g_BuddyBuffer.BuildBuffList 'update buddy buffs list
    Call g_BuddyBuffer.PushBuffs(True, True)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ForceRebuff"
    Resume Fin
End Sub

Public Sub ForceRebuff()
On Error GoTo ErrorHandler

    m_bForceRebuff = True
    Call g_Buffer.BuildBuffList 'update buffs list
    Call g_Engine.FireStartMacro(MODE_REBUFF, g_ui.Main.chCombatType.Selected)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ForceRebuff"
    Resume Fin
End Sub

'#####################################################################################
'#
'#                                PUBLIC
'#
'#####################################################################################

Public Function StartMacro(Optional ByVal iMode As Integer = MODE_NORMAL, Optional ByVal iCombatType As Integer = TYPE_CASTER) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    Dim sReason As String
    
    bRet = True

1    g_Core.TimerInterval = DEFAULT_MACRO_TIMER_INTERVAL
2    m_iCombatType = iCombatType

    'If mAuth.m_PluginEnabled = False Then
    '    PrintMessage g_String(mStrings.e_strUnableToAuth)
    '    bRet = False
    '    GoTo Fin
    'End If
    
    'Security - Check Auth
    'If MustLogOut(sReason) Then
    '    PrintMessage "Unable to start macro : " & sReason
    '    bRet = False
    '    GoTo Fin
    'End If
    
    'remove pause
    If Paused Then TogglePause
        
    'is macro already running?
    If Ticking Then
        If iMode = MODE_REBUFF Then
3            Call g_Buffer.PushBuffs(True, True)   'just push the whole buff list to the buff queue
        End If
    Else
        'make sure we have set all the required macro items first
        If Not ValidMacroItems(m_iCombatType) Then
            bRet = False
            GoTo Fin
        End If
        
        m_iMode = iMode
        
        Select Case m_iMode
            Case MODE_REBUFF
                m_iMode = MODE_REBUFF
4                Call g_Buffer.StartService(REBUFF_FULL, , False)
5                Call GoRebuff
            
            Case Else
                If g_ui.Buffs.chkEnableBuffing.Checked Then
                    Dim iBuffMode As Integer
                    
                    If g_ui.Buffs.chkContinuousBuffing.Checked Then
                        iBuffMode = REBUFF_CONTINUOUS
                    Else
                        iBuffMode = REBUFF_FULL
                    End If
                
6                    Call g_Buffer.StartService(iBuffMode, g_Data.RebuffInterval, True, g_ui.Buffs.chkRebuffOnStart.Checked, g_Data.NumContinuousBuffs)
                    
                    If Valid(g_buffBuddy) Then
7                        Call g_BuddyBuffer.StartService(iBuffMode, g_Data.RebuffInterval, True, g_ui.Buffs.chkRebuffOnStart.Checked, g_Data.NumContinuousBuffs)
                    End If
                End If
                
8                Call GoIdle
                
        End Select
        
9       Call OnMacroStart
        
    End If

Fin:
    StartMacro = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.StartMacro() - " & Err.Description & "  line: " & Erl
    Resume Fin
End Function

Public Sub StopMacro()
    MyDebug "StopMacro - Stopping Engine."
    Call g_Engine.FireStopMacro
    m_bRunning = False
    Call CheckStopNav
    Call g_Buffer.StopService
    Call g_BuddyBuffer.StopService
    If Valid(g_ui) Then Call g_ui.Main.UpdateStatus
End Sub

Public Sub PauseMacro()
    If Active Then
        m_bPaused = True
        PrintMessage "Macro paused."
    End If
End Sub

Public Sub ResumeMacro()
    If Not Active Then
        'If mAuth.m_PluginEnabled = False Then
        '    PrintMessage g_String(mStrings.e_strUnableToAuth)
        '    Exit Sub
        'End If
        m_bPaused = False
        PrintMessage "Macro resumed."
    End If
End Sub

'Make a check list of the required items to macro (wand/bow/shield/weapon, etc)
Public Function ValidMacroItems(Optional ByVal iCombatType = TYPE_NOTYPE) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    bRet = True 'default to ok
    
    If iCombatType = TYPE_NOTYPE Then
        iCombatType = m_iCombatType
    End If
    
    'first update equipment collections
    Call g_Objects.Equipment.Update
    
    If Not CheckMacroItem(ID_WAND, g_Data.Wand, g_Objects.Equipment.Wand) Then
        bRet = False
        GoTo Fin
    End If
    
    'Check Weapon/Shield if MELEE mode
    If iCombatType = TYPE_MELEE Then
        If Not CheckMacroItem(ID_WEAPON, g_Data.Weapon, g_Objects.Equipment.Weapon) Then
            bRet = False
            GoTo Fin
        ElseIf Not CheckMacroItem(ID_SHIELD, g_Data.Shield, g_Objects.Equipment.Shield) Then
            bRet = False
            GoTo Fin
        End If
    ElseIf iCombatType = TYPE_ARCHER Then
        If Not CheckMacroItem(ID_BOW, g_Data.Bow, g_Objects.Equipment.Weapon) Then
            bRet = False
            GoTo Fin
        End If
    End If
    
Fin:
    ValidMacroItems = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.ValidMacroItems - " & Err.Description
    Resume Fin
End Function

Public Sub RequestEquipItem(objItem As acObject)
    If Not Valid(objItem) Then
        PrintWarning "RequestEquipItem : objItem = NULL"
        Exit Sub
    End If
    
    If m_Equip.StartEquip(objItem) Then
        MyDebug "RequestEquipItem : SUBST_EQUIP"
        Call SetSubState(SUBST_EQUIP)
        Call RunState
    End If
End Sub

Public Function ValidCombatState(ByVal iCombatState As eCombatStates)
    
End Function

Public Sub RequestCombatStateChange(ByVal iNewState As eCombatStates)
    If m_CombatMode.ChangeCombatState(iNewState) Then
        Call SetSubState(SUBST_CHANGE_COMBATMODE)
        Call RunState
    End If
End Sub

Public Function ValidState(Optional ByVal CombatType As eMacroCombatType = TYPE_NOTYPE) As Boolean
On Error GoTo Error_Handler

    Dim bRet As Boolean
    bRet = False
    
    If CombatType = TYPE_NOTYPE Then
        CombatType = m_iCombatType
    End If
    
    Select Case CombatType
        Case TYPE_MELEE
            If (Not Valid(g_Data.Weapon)) _
            Or (Not Valid(g_Data.Shield)) Then
                PrintErrorMessage "ValidState : macro Weapon and/or Shield not set."
                GoTo Fin
            End If
        
            'check if valid equipment
            If Not g_Data.Weapon.Equiped Then
                MyDebug "ValidState : Melee Weapon not equipped, equipping now."
                Call RequestEquipItem(g_Data.Weapon)
                GoTo Fin
            ElseIf Not g_Data.Shield.Equiped Then
                MyDebug "ValidState : Shield not equipped, equipping now."
                Call RequestEquipItem(g_Data.Shield)
                GoTo Fin
            Else
                If m_CombatMode.CurrentCombatState <> COMBATSTATE_MELEE And m_Healing.NeedHealing = False Then
                    MyDebug "ValidState: TYPE_MELEE: invalid combat state. Switching now"
                    Call RequestCombatStateChange(COMBATSTATE_MELEE)
                    GoTo Fin
                End If
            End If
        
        Case TYPE_ARCHER
            If Not Valid(g_Data.Bow) Then
                PrintErrorMessage "ValidState : macro Bow not set."
                GoTo Fin
            End If
            
            If Not g_Data.Bow.Equiped Then
               MyDebug "ValidState : Bow not equipped, equipping now."
                Call RequestEquipItem(g_Data.Bow)
                GoTo Fin
            Else
                If m_CombatMode.CurrentCombatState <> COMBATSTATE_ARCHER And m_Healing.NeedHealing = False Then
                    MyDebug "ValidState: TYPE_ARCHER: invalid combat state. Switching now"
                    Call RequestCombatStateChange(COMBATSTATE_ARCHER)
                    GoTo Fin
                End If
            End If
            
        Case Else   'Caster
            If Not Valid(g_Data.Wand) Then
                PrintErrorMessage "ValidState : macro Wand not set."
                GoTo Fin
            End If
            
            If Not g_Data.Wand.Equiped Then
               MyDebug "ValidState : Wand not equipped, equipping now."
                Call RequestEquipItem(g_Data.Wand)
                GoTo Fin
            Else
                If m_CombatMode.CurrentCombatState <> COMBATSTATE_MAGIC Then
                    MyDebug "ValidState: TYPE_CASTER: invalid combat state. Switching now"
                    Call RequestCombatStateChange(COMBATSTATE_MAGIC)
                    GoTo Fin
                End If
            End If
            
    End Select

    'everything went fine
    bRet = True

Fin:
    ValidState = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage "ValidState"
    Resume Fin
End Function


'#####################################################################################
'#
'#                                EVENTS
'#
'#####################################################################################

'A vitals-recharging spell is done being casted
Private Sub OnStateSpellCasted()
On Error GoTo ErrorHandler

    'MyDebug "OnStateSpellCasted: m_iState: " & m_iState
    
    Select Case m_iState
    
        Case ST_HEALING
            If m_Healing.HealingMethod = MET_SPELL Then
                Call m_Healing.OnHealingReady
            End If
        
        Case ST_RESTAM
            If m_Restam.RestamMethod = RESTAM_SPELL Then
                Call m_Restam.OnRestamReady
            End If
    
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnStateSpellCasted() - " & Err.Description
    Resume Fin
End Sub

Private Sub OnStateSpellFailed()
On Error GoTo ErrorHandler
    
    Select Case m_iState
        Case ST_HEALING
            If m_Healing.HealingMethod = MET_SPELL Then
                Call m_Healing.OnHealingReady(False)
            Else
                PrintWarning "OnStateSpellFailed - HealingMethod (" & m_Healing.HealingMethod & ") != MET_SPELL"
            End If
        
        Case ST_RESTAM
            If m_Restam.RestamMethod = RESTAM_SPELL Then
                Call m_Restam.OnRestamReady(False)
            Else
                PrintWarning "OnStateSpellFailed - RestamMethod != RESTAM_SPELL"
            End If
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnStateSpellFailed() - " & Err.Description
    Resume Fin
End Sub

'True if error was handled here
Public Function OnLastActionFailed(Optional ByVal iFailureId As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    bRet = True

    Select Case iFailureId
        Case FAIL_TOO_BUSY, FAIL_BUSY
            ' Do nothing
            MyDebug "-- Fail: Busy --"
            If m_iState <> ST_COMBAT Then
                'Call g_Core.SendKeyHold(g_Keys.KeyJump)
                Call ResetSubState
                'Call g_Core.SendKeyRelease(g_Keys.KeyJump)
            Else
                Call ResetSubState
            End If
            
            
        Case FAIL_NOT_READY, FAIL_UNPREPARED_TO_CAST
            MyDebug "-- Fail: Not Ready/Unprepared --"
            
            If m_tmrBusyDelay.Expired Then
                Call m_tmrBusyDelay.SetNextTime(0.5)    '0.5 secondes pause
            End If
            
            If (g_Spells.LastCastAttempt + 1) >= g_Core.Time Then
                Call g_Spells.OnSpellCastFailed("OnLastActionFailed - Not Ready/Unprepared")
            End If
            
            'Don't force OnReady if in spell casting mode
            If g_Hooks.CombatMode <> COMBATSTATE_MAGIC Or m_Healing.NeedHealing = False Then
                Call OnReady(False)
            End If
            
            'SPK - Added to prevent macro from getting stuck for too long in a weapon switch/combat mode change
            Call ResetSubState
            
        Case FAIL_SPELL_FIZZLE, FAIL_NOT_ENOUGH_MANA
            MyDebug "-- Fail: Fizzle/No Mana --"
            Call g_Spells.OnSpellCastFailed("OnLastActionFailed - Fizzle/Not enough Mana")
        
        Case FAIL_NOT_ENOUGH_COMPONENTS, FAIL_NO_SPELL_TARGET, FAIL_DONT_KNOW_SPELL, FAIL_IMPOSSIBLE_SPELL_PATH
            MyDebug "-- Fail: NoComps/NoSpellTarget/DontKnowSpell/ImpossibleSpellPath --"
            Call g_Spells.OnSpellCastFailed("OnLastActionFailed - No Comps/No Target/No Spell - Unable to continue casting")
        
        Case FAIL_CHARGED_TOO_FAR
            MyDebug "Fail - Charged too far"
            
'        Case FAIL_UNABLE_TO_MOVE_TO_OBJECT, FAIL_ALREADY_IN_USE, FAIL_CANNOT_BE_USED, FAIL_CANT_OPEN_BODY
'            MyDebug "Fail - Unabled to move to object"
'            If m_iState = ST_LOOT Then
'               Call m_Loot.StopLooting("OnLastActionFailed - Unabled to move to object")
'            End If
            
        Case FAIL_TOO_FATIGUED_TO_ATTACK
            MyDebug "Fail - Too fatigued to attack"
            Call m_Combat.StopCombat("OnLastActionFailed - Too fatigued to attack")
            
        Case FAIL_OUT_OF_AMMUNITION
            MyDebug "Fail - Out of arrow"
            If Ticking Then Call OnOutOfArrow
        
        Case FAIL_INCORRECT_TARGET_TYPE
            MyDebug "Fail - Incorrect target type -> Releasing target"
            Call m_Combat.StopCombat("OnLastActionFailed - Incorrect Target Type")
            
        Case Else
            bRet = False
                
    End Select
    
    If m_iState = ST_LOOT Then
        Dim bLootError As Boolean
        bLootError = m_Loot.HandleActionFailure(iFailureId)
        If Not bRet Then bRet = bLootError
    End If

Fin:
    OnLastActionFailed = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.OnLastActionFailed - " & Err.Description
    Resume Fin
End Function

'A spell is done being casted
Public Sub OnSpellCasted(Optional ByVal bSuccessfully As Boolean = True)

    'MyDebug "OnSpellCasted: bSucessfully: " & bSuccessfully & " :: m_iState: " & m_iState

    If bSuccessfully Then
        Call OnStateSpellCasted
    Else
        Call OnStateSpellFailed
    End If

    If m_bRegenVitals Then
       m_bRegenVitals = False
    End If
    
    If m_iState = ST_COMBAT Then
        Call m_Combat.OnLastSpellCasted
    End If
    
    If Not Active Then
        Exit Sub
    End If

    'If TurboMode Then Call RunState
    Call RunState
    
End Sub

'bSuccess = true if last action was completed successfully
'Note : this applys to non spellcasting actions ONLY,
'spell casting actions are handled in OnLastSpellComplete
Public Sub OnReady(Optional ByVal bSuccess As Boolean = True, Optional ByVal bWasCasting As Boolean = False)
On Error GoTo ErrorHandler
        
    If Not m_tmrBusyDelay.Expired Then
        If bWasCasting Then Exit Sub    'spell castings busy handled elsewhere
        bSuccess = False
    End If
    
    locDebug "-- clsMacro.OnReady: bSuccess : " & CStr(bSuccess)
    
    If m_iSubState <> SUBST_NONE Then
        Select Case m_iSubState
            Case SUBST_EQUIP
                If Not bSuccess Then
                    MyDebug "OnReady - Equip Action Failed - Retrying"
                    Call m_Equip.Restart
                End If
            
            Case SUBST_CHANGE_COMBATMODE
                If Not bSuccess And m_Healing.NeedHealing = False Then 'Don't force change if healing.
                    MyDebug "OnReady - ChangeCM Action Failed - Retrying"
                    Call m_CombatMode.Restart
                End If
        End Select
    Else
        Select Case m_iState
            Case ST_HEALING
                If m_Healing.HealingMethod <> MET_SPELL Then
                    'MyDebug "OnReady - OnHealingReady [MET_KIT or MET_POTION]"
                    Call m_Healing.OnHealingReady(bSuccess)
                End If
                
            Case ST_RESTAM
                If m_Restam.RestamMethod <> RESTAM_SPELL Then
                    'MyDebug "OnReady - OnRestamReady"
                    Call m_Restam.OnRestamReady(bSuccess)
                End If
                
            Case ST_FLETCHING
                If bSuccess = False Then
                    locDebug "OnReady - Fletching action failed, restarting."
                    Call m_Fletcher.RestartFletching
                    Exit Sub
                End If
        End Select
    End If
    
    Call RunState

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnReady - " & Err.Description
    Resume Fin
End Sub

Public Sub OnEntityNoLongerAvailable(objEntity As acObject)
On Error GoTo ErrorHandler
    
    If Not Valid(objEntity) Then
        PrintWarning "clsMacro.OnEntityNoLongerAvailable - invalid objEntity"
        GoTo Fin
    End If
    
    ' Death Message Checker
    If Valid(m_Combat.Target) Then
        
        If m_Combat.Target.Guid = objEntity.Guid Then
            'Now it's safe to release target data
            Call m_Combat.StopCombat("clsMacro.OnEntityDeath")
        End If
        
        'Call Macro.Timers(TMR_LOOT_DELAY).SetNextTime(2)    'wait a bit before looting this body, as it may not be ready yet
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnEntityNoLongerAvailable - " & Err.Description
    Resume Fin
End Sub

Public Sub OnEntityDeath(objEntity As acObject)
On Error GoTo ErrorHandler
    
    locDebug "Entity Death : " & objEntity.Name
    Call OnEntityNoLongerAvailable(objEntity)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnEntityDeath - " & Err.Description
    Resume Fin
End Sub

Public Sub setLowManaCheck()
    If g_ui.Macro.chkUseManaCharge.Checked Then
        m_bNeedManaCharge = True
        m_bUsingManaCharge = False
    End If
End Sub

Private Sub goManaCharge()
On Error GoTo ErrorHandler

    Call SetState(ST_MANACHARGE)
    Call m_tmrManaChargeTimer.SetNextTime(8)    'Allow 8 seconds for mana charge to be used
    Call m_tmrBusyDelay.SetNextTime(1)          'One second pause between tries
    
    Call Vitals.idManaStones
   
    If Vitals.findManaCharge Then
        If Valid(g_manaItem) And (g_manaItem.Guid <> -1) Then
            MyDebug "goManaCharge: Using " & g_manaItem.Name
            Call g_Service.UseItemOnSelf(g_manaItem)
            m_bUsingManaCharge = True
        Else
            PrintMessage "Could not find any Mana Charges"
            Call doneManaCharge
        End If
    Else
        PrintMessage "Could not find any Mana Charges"
        Call doneManaCharge
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.goManaCharge - " & Err.Description
    Resume Fin
End Sub

Private Sub UseManaCharge()
On Error GoTo ErrorHandler

    If Not (m_bNeedManaCharge) Then Call doneManaCharge
    
    If Not m_tmrBusyDelay.Expired Then
        Call Vitals.idManaStones
        Exit Sub
    End If

    If Valid(g_manaItem) And (g_manaItem.Guid <> -1) Then
        MyDebug "UseManaCharge: Using " & g_manaItem.Name
        Call g_Service.UseItemOnSelf(g_manaItem)
        m_bUsingManaCharge = True
    Else
        ' Must have used the mana charge already
        Call doneManaCharge
    End If
    
    Call m_tmrBusyDelay.SetNextTime(1)          'One second pause between tries
  
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.UseManaCharge - " & Err.Description
    Resume Fin
End Sub

Public Sub doneManaCharge()
On Error GoTo ErrorHandler
    
    MyDebug "doneManaCharge: all done!"
    Set g_manaItem = Nothing
    m_bNeedManaCharge = False
    m_bUsingManaCharge = False
    Call m_tmrManaChargeTimer.Reset
    Call GoIdle
    Call Vitals.idManaStones

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.doneManaCharge - " & Err.Description
    Resume Fin
End Sub

Private Sub ReloadArrows()
On Error GoTo ErrorHandler

    Dim NumArrowsInInventory As Integer
    Dim ArrowName As String
    Dim ArrowPileGUID As Long
    Dim NumArrowHeads As Integer
    Dim NumShafts As Integer
    Dim objArrow As acObject

    ArrowName = g_Data.Arrows
    PrintMessage "Out of " & ArrowName & ". Looking up for some in inventory..."
    NumArrowsInInventory = g_Objects.Items.InvCntByName(ArrowName, True)
    
    If NumArrowsInInventory > 0 Then
        PrintMessage "Found a total of " & NumArrowsInInventory & " " & ArrowName & "(s) in inventory. Using first pile."

        Set objArrow = g_Objects.Items.InvFindByName(ArrowName)
        If objArrow.Guid <> -1 Then
            Call RequestEquipItem(objArrow)
        Else
            PrintErrorMessage "OnOutOfArrow : unable to find the arrow object in inventory"
            GoTo Fin
        End If
    Else
        'Fletch if needed and enabled
        If g_ui.Macro.chkEnableFletching.Checked Then
            'look up for arrow heads
            NumArrowHeads = g_Objects.Items.InvCntByName(g_Data.ArrowHead, True)
            MyDebug "...NumArrowHeads = " & NumArrowHeads

            'look up for shafts
            NumShafts = g_Objects.Items.InvCntByName(g_Data.ArrowShaft, True)
            MyDebug "...NumShafts = " & NumShafts

            If (NumArrowHeads > 0) And (NumShafts > 0) Then
                MyDebug "Found enough ArrowHeads/Shafts, starting fletching..."
                Call GoFletching
                GoTo Fin
            Else
                MyDebug "Out of fletching components..."
            End If
        End If

        If g_ui.Macro.chkLogoutOnArrows.Checked Then
            'FIXME: add options for melee mode swap or logging
            PrintMessage "No more arrows/fletching components in inventory. Turning macro off."
            LogEvent "Macro out of " & ArrowName & " and fletching components."
            MyDebug "Fletching - out of Arrows and Components !"
            Call StopMacro
            Call g_Service.Logout("No more arrows/fletching components in inventory")
        End If
    
    End If
    
Fin:
    m_bNeedArrows = False   'reset flag once this function has been called
    Set objArrow = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.OnOutOfArrow - " & Err.Description
    Resume Fin
End Sub

Public Sub OnOutOfArrow()
    m_bNeedArrows = True
    Call m_tmrBusyDelay.SetNextTime(2)  'give some delay to switch back to peace mode after out of ammo
    Call GoIdle
End Sub

'#####################################################################################
'#
'#                                UTILS
'#
'#####################################################################################

'Check logout conditions
Public Function MustLogOut(Optional ByRef sReasonOut As String) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    bRet = False
    
    If (Not bRet) And g_ui.Macro.chkLogoutOnDeath.Checked And m_bDied Then
        sReasonOut = "Macro died"
        bRet = True
    End If
    
    If (Not bRet) And g_ui.Macro.chkLogoutOnTapers.Checked Then
        Dim iTaperCount As Integer
        iTaperCount = g_Objects.Items.InvCntByName(STR_ITEM_TAPER)
        If iTaperCount < 50 Then
            sReasonOut = "Running out of Prismatic Tapers (" & iTaperCount & ")"
            bRet = True
            ' Never force logout if we need healing!
            If g_Macro.NeedHealing("clsMacro.MustLogOut") Then bRet = False
        End If
    End If
    
    If (Not bRet) And g_ui.Macro.chkLogoutOnKits.Checked Then
        If g_ui.Macro.chkUseHealingKits.Checked Then
            If Not Valid(Vitals.findHealItem) Then
                sReasonOut = "Out of Healing Items"
                bRet = True
            End If
        End If
    End If
    
    If (Not bRet) And g_ui.Macro.chkLogoutOnScarabs.Checked Then
        If g_Objects.Items.InvCntByName(STR_ITEM_PLAT) <= 0 Then
            sReasonOut = "Running out of " & STR_ITEM_PLAT
            bRet = True
            ' Never force logout if we need healing!
            If g_Macro.NeedHealing("clsMacro.MustLogOut") Then bRet = False
        ElseIf g_Objects.Items.InvCntByName(STR_ITEM_PYREAL_SCARAB) <= 0 Then
            sReasonOut = "Running out of " & STR_ITEM_PYREAL_SCARAB
            bRet = True
            ' Never force logout if we need healing!
            If g_Macro.NeedHealing("clsMacro.MustLogOut") Then bRet = False
        End If
    End If
    
    If (Not bRet) And g_ui.Macro.chkUseRings.Checked Then
        If g_Objects.Items.InvCntByName(STR_ITEM_DIAMOND_SCARAB) <= 0 Then
            PrintMessage "NOTICE: Running out of Diamond Scarabs!"
            g_ui.Macro.chkUseRings.Checked = False
        End If
    End If
    
    If (g_AntiBan.m_bLogOut) Then
        Call g_Service.Logout(g_AntiBan.m_sReason)
        bRet = True
    End If
    
Fin:
    MustLogOut = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacro.MustLogOut - " & Err.Description
    Resume Fin
End Function

Private Sub SetElapsedTime(ByVal dTime As Double)
On Error GoTo ErrorHandler

    If (dTime <= 0) Then
        m_dElapsedTime = 1
    Else
        m_dElapsedTime = dTime
    End If
    
    m_sTimeSpent = myFormatTime(m_dTimeSpent, TF_LETTERS)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.SetElapsedTime - " & Err.Description
    Resume Fin
End Sub

Private Sub UpdateNextRebuffDisplay(Optional ByVal dTime As Double = -1)
On Error GoTo ErrorHandler

1    If (dTime = -1) And Valid(g_Buffer) Then
2        dTime = g_Buffer.NextRebuff.RemainingTime
    End If
     
3    If dTime < 0 Then
         dTime = 0
     End If
     
     If dTime > 0 Then
          m_sNextRebuff = myFormatTime(dTime, TF_LETTERS)
     Else
          m_sNextRebuff = "Unknown"
     End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.UpdateNextRebuffDisplay - " & Err.Description & " - line " & Erl
    Resume Fin
End Sub

Public Function GetSecondsUntilLevel() As Double
On Error GoTo Error_Handler

    Dim fXpToNextLvl As Double, fSecsRemainingToLvl As Double

    If m_dXpHour <= 0 Then
        GetSecondsUntilLevel = 0
    Else
        fXpToNextLvl = g_ds.XpTracker.XPToNextLevel
        GetSecondsUntilLevel = (CDbl(fXpToNextLvl) * CDbl(3600)) / m_dXpHour
    End If
        
Fin:
    Exit Function
Error_Handler:
    PrintErrorMessage "clsMacro.GetSecondsUntilLevel"
    Resume Fin
End Function

Private Sub UpdateTimeUntilLevelDisplay(Optional ByVal dTime As Double = -1)
On Error GoTo ErrorHandler

    If (dTime = -1) Then dTime = GetSecondsUntilLevel
    If dTime <= 0 Then dTime = 0
    m_sTimeUntilLevel = myFormatTime(dTime, TF_LETTERS)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.UpdateTimeUntilLevelDisplay"
    Resume Fin
End Sub

Private Sub UpdateStats()
On Error GoTo ErrorHandler
    
    If (m_dXpAtStart = 0) Then
        m_dXpAtStart = g_ds.XpTracker.TotalXp
    End If
    
    m_dTotalXp = g_ds.XpTracker.TotalXp
    m_dXpEarned = m_dTotalXp - m_dXpAtStart
    
    m_dTimeSpent = m_dElapsedTime - m_dLastResetTime
    
    m_sTimeSpent = myFormatTime(m_dTimeSpent, TF_LETTERS)

    If m_dTimeSpent <= 0 Then
        m_dXpHour = 0
        m_dTapersHour = 0
        m_dArrowsHour = 0
        m_dPlatsHour = 0
        m_dFiveMinXPAvg = 0
    Else
        m_dXpHour = (m_dXpEarned / m_dTimeSpent) * CDbl(3600)
        m_dTapersHour = (m_dTapersCount / m_dTimeSpent) * CDbl(3600)
        m_dPlatsHour = (m_dPlatsCount / m_dTimeSpent) * CDbl(3600)
        m_dArrowsHour = (m_dArrowsCount / m_dTimeSpent) * CDbl(3600)
    End If
       
    Call UpdateStatsStrings
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.UpdateStats - " & Err.Description
    Resume Fin
End Sub

Public Sub ResetXpStats()
On Error GoTo ErrorHandler

    m_dTotalXp = g_ds.XpTracker.TotalXp
    m_dXpAtStart = m_dTotalXp
    m_dXpEarned = 0
    m_dXpHour = 0
    m_dTapersCount = 0
    m_dTapersHour = 0
    m_dPlatsCount = 0
    m_dPlatsHour = 0
    m_dArrowsCount = 0
    m_dArrowsHour = 0
    m_dLastResetTime = m_dElapsedTime
    g_TotalKilled = 0
    g_TotalLooted = 0
    m_dFiveMinXPAvg = 0
    m_dFiveMinXpHour = 0
    Call UpdateStatsStrings

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.ResetXpStats - " & Err.Description
    Resume Fin
End Sub

Private Sub UpdateStatsStrings()
On Error GoTo ErrorHandler
    
    m_sXpHour = FormatXp(m_dXpHour)
    m_sXpMinute = FormatXp(m_dFiveMinXpHour)
    m_sXpEarned = FormatXp(m_dXpEarned)
    m_sTotalXp = FormatXp(m_dTotalXp)
    m_sTapersHour = FormatXp(m_dTapersHour, False)
    m_sPlatsHour = FormatXp(m_dPlatsHour, False)
    m_sArrowsHour = FormatXp(m_dArrowsHour, False)
    Call UpdateTimeUntilLevelDisplay
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.UpdateStatsStrings - " & Err.Description
    Resume Fin
End Sub


Public Function GetStateString() As String
On Error GoTo ErrorHandler
    
    Dim sRet As String
    
    If Not Ticking Then
        sRet = "Off"
    ElseIf Paused Then
        sRet = "Paused"
    Else
        Select Case m_iState
            Case ST_IDLE
                sRet = "Idle"
                
            Case ST_REBUFF
                sRet = "Rebuff"
                
            Case ST_BUDDYREBUFF
                sRet = "Buddy Rebuff"
                            
            Case ST_COMBAT
                sRet = "Fighting"
                If Valid(m_Combat.Target) Then
                    sRet = sRet & " " & m_Combat.Target.Name
                End If
                
            Case ST_HEALING
                sRet = "Healing"
                
            Case ST_RESTAM
                sRet = "Restaming " & m_Restam.GetCurrentStateString
            
            Case ST_FLETCHING
                sRet = "Fletching (" & m_Fletcher.GetStateString & ")"
                
            Case ST_STICK_NAV
                sRet = "Moving (sticky)"
                
            Case ST_NAV
                sRet = "Moving"
                
            Case ST_FELLOW_CMD
                sRet = "Fellow Cmd"
                
            Case ST_LOOT
                sRet = "Looting (" & m_Loot.GetStateString & ")"
                
            Case ST_SALVAGE
                sRet = "Salvaging (" & m_Salvager.GetStateString & ")"
                
            Case ST_BUYSELL
                sRet = "Buy/Sell"
            
            Case ST_MANACHARGE
                sRet = "Mana Charge"
            
            Case Else
                sRet = "Unknown Macro State " & m_iState
        End Select
        
        If m_iSubState <> SUBST_NONE Then
            Select Case m_iState
                Case SUBST_CHANGE_COMBATMODE
                    sRet = sRet & " (Change CombatMode)"
                Case SUBST_EQUIP
                    sRet = sRet & " (Equip item)"
            End Select
        End If
        
    End If
    GetStateString = sRet

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsMacro.GetStateString - " & Err.Description
    Resume Fin
End Function

Public Sub incLoot()
    Call m_Loot.incNavHold
End Sub

Public Sub SetState(ByVal iState As Integer)
    Call ResetSubState
    m_iState = iState
End Sub

Private Sub SetSubState(ByVal iSubState As Integer)
    m_iSubState = iSubState
    If TurboMode Then Call RunState
    Call RunState
End Sub

Private Sub ResetSubState()
    m_iSubState = SUBST_NONE
End Sub

'Used to tell the macro something has changed in the environment and we can do a nearby-target scan right away
'Called in ObjectCreate / ObjectMoved
Public Sub ResetTargetScanTimer()
    Call m_tmrNextTargetScan.Reset
End Sub

Private Sub m_tmrManaChargeTimer_OnTimeout()
    If (m_bUsingManaCharge) Then
        PrintErrorMessage "ManaCharge Timeout. Failed to use mana charge"
    End If
    m_bNeedManaCharge = False
    m_bUsingManaCharge = False
    Set g_manaItem = Nothing
    Call GoIdle
End Sub

Private Sub m_tmrCastSpell_OnTimeout()
    m_bOKToCast = True
End Sub

Private Sub m_tmr5MinTimer_OnTimeout()
On Error GoTo ErrorHandler
  
    Dim prevCounter As Long
    
    prevCounter = FiveMinXPCounter.Position
    
    If (FiveMinXPCounter.startXP(FiveMinXPCounter.Position) = 0) Then
        FiveMinXPCounter.XP(FiveMinXPCounter.Position) = 0
    Else
        MyDebug "5Min Timer: pos: " & FiveMinXPCounter.Position & " T: " & g_ds.XpTracker.TotalXp & " S: " & FiveMinXPCounter.startXP(FiveMinXPCounter.Position)
        ' Calculate the diff between the Current XP and XP from 5 mins ago
        FiveMinXPCounter.XP(FiveMinXPCounter.Position) = g_ds.XpTracker.TotalXp
    End If
  
    MyDebug "5Min Timer: pos: " & FiveMinXPCounter.Position & " : " & FiveMinXPCounter.XP(FiveMinXPCounter.Position)
  
    prevCounter = FiveMinXPCounter.Position
  
    m_dFiveMinXPAvg = FiveMinXPCounter.XP(FiveMinXPCounter.Position) - FiveMinXPCounter.startXP(FiveMinXPCounter.Position)
    'm_dFiveMinXPAvg = m_dFiveMinXPAvg + FiveMinXPCounter.XP(prevCounter) - FiveMinXPCounter.XP(FiveMinXPCounter.Position)
    'm_dFiveMinXPAvg = CDbl(m_dFiveMinXPAvg + (FiveMinXPCounter.XP(FiveMinXPCounter.Position) * 12) / 2)
    m_dFiveMinXpHour = CDbl(m_dFiveMinXPAvg * 12)
    'm_dFiveMinXpHour = CDbl(FiveMinXPCounter.XP(FiveMinXPCounter.Position) * 12)
    
    'Move to next counter position
    FiveMinXPCounter.Position = (FiveMinXPCounter.Position + 1) Mod (FiveMinXPCounter.Count + 1)
    
    FiveMinXPCounter.XP(FiveMinXPCounter.Position) = 0
    FiveMinXPCounter.startXP(FiveMinXPCounter.Position) = g_ds.XpTracker.TotalXp

    Call m_tmr5MinTimer.SetNextTime(300)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacro.m_tmr5Min - " & Err.Description
    Resume Fin
End Sub

'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = False)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsMacro : " & GetStateString & "] " & DebugMsg, bSilent)
    End If
End Sub

