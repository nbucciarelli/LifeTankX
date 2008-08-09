VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacroCombat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Combat State Engine                               [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False
Private Const SECURE_TIMER = 5  'seconds

Public Enum eSubStates
    ST_COMBAT_NEXTTHINK         'Prepare the next attack action to perform
    ST_COMBAT_ATTACKING         'Attack is occuring
    ST_COMBAT_END               'Combat is over
End Enum

Private m_State As eSubStates           'current combat state
Private m_objTarget As acObject         'our current target
Private m_lastTargetName As String      'last target name
Private m_lLastTargetGUID As Long
Private m_lastSpellTarget As Long       'GUID of last spell's target
Private m_bAttackKeyPushed As Boolean
Private m_tMark As CD3DObj
Private m_Action As eActionType
Private m_tmrAttackCheck As clsTimer
Private m_tmrSecure As clsTimer
Attribute m_tmrSecure.VB_VarHelpID = -1

Public tmrDanger As clsTimer
Public m_tmrBusyCombat As clsTimer        'Expires when the macro no longer in combat animation

Public Event OnStopCombat()


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Set m_tmrAttackCheck = CreateTimer
    Set m_tmrSecure = CreateTimer
    Set tmrDanger = CreateTimer
    Set m_tmrBusyCombat = CreateTimer
    Call Reset
End Sub

Public Sub Reset()
    Set m_objTarget = Nothing
    Call m_tmrAttackCheck.Reset
    Call m_tmrSecure.Reset
    Call tmrDanger.Reset
    Call m_tmrBusyCombat.Reset
    m_State = ST_COMBAT_NEXTTHINK
    m_Action = ACT_NONE
    m_bAttackKeyPushed = False
    m_lLastTargetGUID = 0
End Sub

Private Sub Class_Terminate()
    Set m_objTarget = Nothing
    Set m_tmrAttackCheck = Nothing
    Set m_tmrSecure = Nothing
    Set tmrDanger = Nothing
    Set m_tmrBusyCombat = Nothing
    Set m_tMark = Nothing
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'current attack substate
Public Property Get State() As Integer
    State = m_State
End Property

'current target
Public Property Get Target() As acObject
    Set Target = m_objTarget
End Property

Public Property Get lastTargetName() As String
    lastTargetName = m_lastTargetName
End Property

Public Property Get CastAction() As eActionType
    CastAction = m_Action
End Property

Public Property Get AttackMoveCompleted() As Boolean
    AttackMoveCompleted = m_tmrBusyCombat.Expired
End Property

Public Sub updateSecureTimer(ByVal aTime As Long)
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER + aTime)
End Sub
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                 Public Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'******************************************
' EngageCombat (Target)
'
' Engage combat with objTarget
'******************************************
Public Function EngageCombat(objTarget As acObject, Optional sSource As String = "") As Boolean
On Error GoTo ErrorHandler
    
    locDebug "EngageCombat - Src: " & sSource
    
    'default ret val
    EngageCombat = False
    
    If Not Valid(objTarget) Then
        PrintWarning "clsMacroCombat.EngageCombat : invalid objTarget"
        Exit Function
    End If
    
    'Set new target
    Call SetTarget(objTarget)
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    
    'Prepare attack
    Call SetSubState(ST_COMBAT_NEXTTHINK)
    
    'Return ok
    EngageCombat = True
    
Fin:
    Exit Function
ErrorHandler:
    EngageCombat = False
    PrintErrorMessage "clsMacroCombat.EngageCombat - " & Err.Description
    Resume Fin
End Function


'******************************************
' EndCombat
'
' Stops combat, and release Combat state
'******************************************
Public Sub StopCombat(Optional sSource As String = "")
On Error GoTo ErrorHandler
    
    locDebug "StopCombat - Src: " & sSource
    
    'DS Filter hack to deal with no getting OnDelete messages
    'If Valid(m_objTarget) Then
    '    Dim i As Integer
    '    i = m_objTarget.UserData(INT_DELETE) + 1
    '    Call m_objTarget.SetUserData(INT_DELETE, i)
    '    'MyDebug "clsMacroCombat.StopCombat: INT_DELETE: (" & i & ") " & m_objTarget.Guid & " : " & m_objTarget.Name
    '    If (m_objTarget.UserData(INT_DELETE) > 3) Then
    '        m_objTarget.canDelete = True
    '        MyDebug "clsMacroCombat.StopCombat: canDelete: " & m_objTarget.Guid & " : " & m_objTarget.Name
    '        PrintErrorMessage "StopCombat: fighting non-existant creature, setting canDelete"
    '    End If
    'End If
    
    Call ReleaseTarget("StopCombat: " & sSource)
    Call m_tmrSecure.Reset
    Call m_tmrAttackCheck.Reset
    If g_Spells.Casting Then
            locDebug "clsMacroCombat.StopCombat, calling OnSpellCastComplete"
            Call g_Spells.OnSpellCastComplete(True, "StopCombat")
    End If

    RaiseEvent OnStopCombat

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacroCombat.StopCombat - " & Err.Description
    Resume Fin
End Sub


'******************************************
' RunState
'
' Combat State Engine
'******************************************
Public Sub RunState()
On Error GoTo ErrorMessage

    If m_tmrSecure.Expired Then
        'PrintWarning "Busy Timeout - reseting combat flags"
        MyDebug "clsMacroCombat: m_tmrSecure.Expired - reseting combat, calling StopCombat"
        Call StopCombat("clsMacroCombat.m_tmrSecure.Expired")
        Exit Sub
    End If

    If g_Macro.MustHeal Then
        Call g_Macro.GoIdle
        'RaiseEvent OnStopCombat

        ' Do not call StopCombat, as we will loose the current Target
        'Call StopCombat("Must Heal")
        Exit Sub
    End If
    
    If (g_Macro.CombatType = TYPE_MELEE) And g_Nav.CheckStickyRange(g_ui.Macro.txtMeleeRadius.Text) Then
        Call StopCombat("Sticky Range exceeded, go IDLE")
        Exit Sub
    End If

    ' *** FIXME :: need to add a high priority buff UI check box first ***
    'If we are a spell caster, then can buff while killing
    'If IsCaster Then
    '    'check if we need to buff
    '    If (g_ui.Buffs.chkEnableBuffing.Checked) And (g_Buffer.BuffQueue.Count > 0) Then
    '        Call StopCombat("Caster needs to buff")
    '        Exit Sub
    '    End If
    'End If
    
    Select Case m_State

        'Called before the each new attack
        Case ST_COMBAT_NEXTTHINK
            If Not CurrentTargetValid Then
                locDebug "ST_COMBAT_NEXTTHINK : current target not valid"
                Call StopCombat("ST_COMBAT_NEXTTHINK - Invalid Target")
            Else    'ready to kill
                If IsMelee Then
                    'make sure target is selected
                    Call MeleeCheckSelection
                    
                    'Perform attack (mage/melee/archer)
                    Call FireAttack

                ElseIf IsCaster Then
                    'make sure we can cast
                    If g_Macro.isReadyToCast(TurboMode) Then
                        locDebug "clsMacroCombat.RunState: isReady to cast, FireAttack()"
                        Call FireAttack
                    Else
                        locDebug "clsMacroCombat.RunState: IsCaster: not isReadyToCast()"
                    End If
                End If
            End If
        
        'While combat animation is running (mostly extra selection validity checks)
        Case ST_COMBAT_ATTACKING
        
            'check if stuck in attack combat state for too long
            If IsMelee Then
                'make sure target is selected
                Call MeleeCheckSelection
                'check if we need to push the attack key again if it didnt work the first time
                If m_bAttackKeyPushed And m_tmrAttackCheck.Expired Then
                    locDebug "clsMacroCombat.RunState: ST_COMBAT_ATTACK:Melee: Previous Attack Key Push failed, retrying now"
                    Call FireAttack
                End If
            ElseIf IsCaster Then
                'Keep trying to cast 'till we get the spell words
                If Not g_Spells.Casting Then
                    locDebug "clsMacroCombat.RunState: ST_COMBAT_ATTACK:Caster: Pounding the spell"
                    Call FireAttack
                End If
            End If
            
    End Select
    
Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroCombat.RunState - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               External Events
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'NOTE: this is a particular case where g_macro.State = ST_COMBAT isnt a precondition
Public Sub OnMeleeAttackComplete()
    If g_Macro.State = ST_COMBAT Then
        Call m_tmrSecure.SetNextTime(SECURE_TIMER)
        If (g_Macro.CombatType = TYPE_ARCHER) Then
            'Only needed for Archers/TW to count ammo
            'Melee's attack lots and lots before critter is dead
            Call OnLastAttackComplete
        End If
    Else 'after killing a monster, we may leave ST_COMBAT before OnMeleeAttackComplete event fires
        'locDebug "OnMeleeAttackComplete - Not ST_COMBAT but resetting Combat timers"
        Call ResetMeleeAttack
    End If
End Sub

'StateEngine is active (enabled & not paused)
Public Sub OnLastSpellCasted()
    If IsCaster Then
        Call m_tmrSecure.SetNextTime(SECURE_TIMER)
        Call OnLastAttackComplete
    End If
End Sub

Public Sub MageImperilLanded()
    If Not Valid(m_objTarget) Then
        MyDebug "clsMacroCombat - Imperil Landing : invalid m_objTarget"
        Exit Sub
    End If
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    locDebug "MageImperilLanded: Monster Imperiled"
    Call m_objTarget.SetUserData(INT_MISSCOUNT, 0)
    m_objTarget.Imperiled = True
End Sub

Public Sub MageVulnLanded()
    If Not Valid(m_objTarget) Then
        MyDebug "clsMacroCombat - Vuln Landing : invalid m_objTarget"
        Exit Sub
    End If
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    locDebug "MageVulnLanded: Monster Vulned."
    Call m_objTarget.SetUserData(INT_MISSCOUNT, 0)
    Call m_objTarget.AddVuln(g_Spells.CurrentDamage)
End Sub

Public Sub MageYieldLanded()
    If Not Valid(m_objTarget) Then
        MyDebug "clsMacroCombat - Yield Landing : invalid m_objTarget"
        Exit Sub
    End If
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    'If (Not m_objTarget.Yielded) And (g_ui.Macro.chkYield.Checked) Then
     locDebug "MageYieldLanded: Monster Yielded."
     Call m_objTarget.SetUserData(INT_MISSCOUNT, 0)
     m_objTarget.Yielded = True
    'End If
End Sub

Public Sub MageSpellGotResisted()
    If Not Valid(m_objTarget) Then
        MyDebug "clsMacroCombat - MageSpellGotResisted : invalid m_objTarget"
        Exit Sub
    End If
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    Call m_objTarget.SetUserData(INT_MISSCOUNT, 0)
    If m_Action = ACT_YIELDMONSTER Then
        Call m_objTarget.SetUserData(INT_YIELD_TRYS_LEFT, m_objTarget.UserData(INT_YIELD_TRYS_LEFT) - 1)
        locDebug "MagicYield: monster resisted yield, " & m_objTarget.UserData(INT_YIELD_TRYS_LEFT) & " trys left"
    End If
End Sub


'**************************************
' SetTarget
'
' Configure the combat for the new target
'
'**************************************
Public Function SetTarget(newTarget As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not Valid(newTarget) Then
        MyDebug "clsMacroCombat.SetTarget: invalid objTarget"
        GoTo Fin
    End If

    
    'if we already have a target, release it
    If Valid(m_objTarget) Then
        m_lastTargetName = m_objTarget.Name
        Call ReleaseTarget("From SetTarget")
    Else
        m_lastTargetName = newTarget.Name
    End If
    
    MyDebug "SetTarget: setting target to " & newTarget.Name
    
    'set reference to object
    Set m_objTarget = newTarget
    
    'Select the monster for when we fire attack key if in non casting mode
    If IsMelee Then
        Call g_Service.SelectObject(m_objTarget)
    End If
    
    Call SetCurrentEquip
    
    'return true
    bRet = True
    
    If g_ui.Options.chk3DTarget.Checked Then
        'Put a 3D arrow above the target
        Set m_tMark = g_D3D.m_d3ds.MarkObjectWithShape(m_objTarget.Guid, D3DService.eShape.eVArrow, &H77BB0055) ' FFBB0000 :: 55667788
        Call m_tMark.AnchorToObject(m_objTarget.Guid, 0.9, 0, 0, 0)
        Call m_tMark.SetScale(0.3)
        m_tMark.autoscale = False
        locDebug "SetTarget: D3D MarkObject " & m_objTarget.Guid
    End If

Fin:
    SetTarget = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "SetTarget - " & Err.Description
    Resume Fin
End Function

' Set the target equipment we should be using
Public Sub SetCurrentEquip()
On Error GoTo ErrorHandler

    If Valid(m_objTarget) Then
        If m_objTarget.Vulnerability <> -1 Then
            Set g_currentEquip = g_Data.GetWeaponType(m_objTarget.Vulnerability)
            If Valid(g_currentEquip) Then
                locDebug "SetCurrentEquip: Setting g_CurrentEquip to: " & g_currentEquip.Name & " for damage type: " & m_objTarget.Vulnerability
                'If g_currentEquip.itemType = ITEM_MISSILE_WEAPON Then
                If g_Macro.CombatType = TYPE_ARCHER Then
                    'Also need to find and equip Arrows
                    locDebug "SetCurrentEquip: TYPE_ARCHER, so checking for correct ammo"
                    Call g_Data.GetArrowType
                End If
            End If
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SetCurrentEquip - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Private Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub SetSubState(stateId As eSubStates)
    m_State = stateId
End Sub

'**************************************
' ReleaseTarget
'
' Resets the current target datas
'**************************************
Private Sub ReleaseTarget(Optional sSource As String = "")
On Error GoTo ErrorHandler

    MyDebug "ReleaseTarget: " & sSource
    
    'Remember current target
    If Valid(m_objTarget) Then
        m_lLastTargetGUID = m_objTarget.Guid
        MyDebug "ReleaseTarget: (" & m_objTarget.Guid & "):: " & m_objTarget.Name
    End If
    
    'Release current target reference
    Set m_objTarget = Nothing

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsCombat.ReleaseTarget - " & Err.Description
    Resume Fin
End Sub

'**************************************
' ResetMeleeAttack
'
' Called after the attack move is completed
' to reset some timers
'
'**************************************
Private Sub ResetMeleeAttack()

    m_bAttackKeyPushed = False
    Call m_tmrAttackCheck.Reset
    
    'SPK - Added 8/06/04
    Call m_tmrBusyCombat.Reset 'Tell the macro we really be done swinging

End Sub

'******************************************
' OnLastAttackComplete
'
' Must be called when a melee/archer/mage attack
' just completed/landed (successfuly or not)
'******************************************
Private Sub OnLastAttackComplete()
On Error GoTo ErrorHandler
    
    locDebug "OnLastAttackComplete"
    
    'Increment Arrow counter for Archers
    If (g_Macro.CombatType = TYPE_ARCHER) Then
        Call g_Macro.ArrowsCount(1)
    End If
    
    'release attack key flag for melee mode
    Call ResetMeleeAttack
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    
    'Only attack again if we don't have a better target to acquire
    If Valid(m_objTarget) Then
    
        ' Check to see if we should Blacklist this critter
        If (g_ui.Macro.chkEnableBlacklist.Checked Or g_ui.Macro.chkEnableArcherBlacklist.Checked) _
        And (m_objTarget.UserData(INT_MISSCOUNT) > BLACKLIST_MAXCOUNT) Then
            Call m_objTarget.SetUserData(INT_BLIST_TIME, g_Core.Time + BLACKLIST_TIME)
            Call m_objTarget.SetUserData(INT_MISSCOUNT, 0)
            MyDebug "Blacklisted: " & m_objTarget.Name
            PrintMessage "Blacklisted: " & m_objTarget.Name
            Dim bCount As Integer
            bCount = m_objTarget.UserData(INT_BLISTCOUNT) + 1
            Call m_objTarget.SetUserData(INT_BLISTCOUNT, bCount)
            ' Check for perma blacklist
            If (m_objTarget.UserData(INT_BLISTCOUNT) > BLACKLIST_MAXCOUNT) Then
                Call m_objTarget.SetUserData(INT_BLIST_TIME, g_Core.Time + 600)
                MyDebug "Blacklisted for 5 mins: " & m_objTarget.Name
                PrintMessage "Blacklisted  for 5 mins: " & m_objTarget.Name
            End If
        End If
        
        If EngagingDangerousTarget Then
            Exit Sub
        ElseIf EngagingBetterTarget Then
            Exit Sub
        Else
            Call SetSubState(ST_COMBAT_NEXTTHINK)
        End If
    Else
        Call StopCombat("OnLastAttackComplete (objTarget is nothing)")
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsCombat.OnLastAttackComplete - " & Err.Description
    Resume Fin
End Sub

'**************************************
' CurrentTargetValid
'
' Returns true if current target is valid
'**************************************
Private Function CurrentTargetValid() As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    CurrentTargetValid = False
    
    If Not IsValidTarget(m_objTarget, False) Then
        locDebug "CurrentTargetValid - invalid Target"
        bRet = False
    'ElseIf Not TargetCanBeReached(m_objTarget) Then
    '    locDebug "CurrentTargetValid - " & m_objTarget.Name & " can't be reached"
    '    bRet = False
    Else
        bRet = True
    End If
    
Fin:
    CurrentTargetValid = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsCombat.CurrentTargetValid - " & Err.Description
    Resume Fin
End Function

'**************************************
' FireAttack
'
' Triggers the mage or melee attack,
' depending of the current macro combat mode
'**************************************
Private Sub FireAttack()
On Error GoTo ErrorHandler

    Dim bAttackFired As Boolean
    
    'fire the attack
    If IsMelee Then
        bAttackFired = FireAttackKey
    ElseIf IsCaster Then
        bAttackFired = FireMageAttack
    End If
    
    If bAttackFired Then
        'The busyCombat timer only applies to melee/archer mode - For mage combat, the completion
        'of the spell casting is what determines when the last attack move is over
        If IsMelee Then
            Call m_tmrBusyCombat.SetNextTime(3) 'SPK added 08/06/04 - The timer is reset when the attack move compeletes in MeleeAttackComplete
        End If
        'Set the new state
        Call SetSubState(ST_COMBAT_ATTACKING)
    Else
        'PrintWarning "clsMacroCombat - FireAttack : attack not fired"
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsCombat.FireAttack - " & Err.Description
    Resume Fin
End Sub

Public Sub PushAttackKey()
On Error GoTo ErrorHandler
    
    Dim lKeyCode As Long

    Select Case g_ui.Macro.chAttackHeight.Selected
        Case ATK_HIGH
            lKeyCode = g_Keys.KeyAttackHigh
            
        Case ATK_MEDIUM
            lKeyCode = g_Keys.KeyAttackMedium
            
        Case ATK_LOW
            lKeyCode = g_Keys.KeyAttackLow
            
        Case Else
            PrintWarning "myAttackKey - Unsupported attack height #" & g_ui.Macro.chAttackHeight.Selected
            lKeyCode = g_Keys.KeyAttackLow
    End Select
    
    Call g_Core.SendKey(lKeyCode)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacroCombat.PushAttackKey - " & Err.Description
    GoTo Fin
End Sub


'**************************************
' FireAttackKey
'
' Pushs the melee/archer attack key
'**************************************
Private Function FireAttackKey() As Boolean

    If Not CurrentTargetValid Then
        locDebug "WARNING - FireAttackKey : Invalid Target. Leaving combat state."
        Call StopCombat("FireAttackKey, Invalid Target")
        FireAttackKey = False
        Exit Function
    End If

'    locDebug "FireAttackKey()"
    If TargetCanBeReached(m_objTarget) Then
        Call PushAttackKey
        Call m_tmrAttackCheck.SetNextTime(3)
        m_bAttackKeyPushed = True
        FireAttackKey = True
    End If
    
End Function


'**************************************
' FireMageAttack
'
' Casts the mage's next attack
'**************************************
Private Function FireMageAttack() As Boolean
On Error GoTo Error_Handler
    Dim i As Integer
    Dim bRet As Boolean
    Dim longRange As Single
    
    bRet = True
    
    locDebug "clsMacroCombat.FireMageAttack - start"
    
    If Not CurrentTargetValid Then
        locDebug "WARNING - FireMageAttack : Invalid Target. Leaving combat state."
        Call StopCombat("FireMageAttack")
        bRet = False
        GoTo Fin
    End If
    
    If (g_ui.Macro.txtVulnRange.Text > g_ui.Macro.txtMageRadius.Text) Then
        longRange = g_ui.Macro.txtVulnRange.Text
    Else
        longRange = g_ui.Macro.txtMageRadius.Text
    End If
    
    
    'Check if we need to yield
    If (g_ui.Macro.chkYield.Checked _
        And Not m_objTarget.Yielded _
        And (m_objTarget.UserData(INT_YIELD_TRYS_LEFT) > 0) _
        And (m_objTarget.UserData(B_CAN_BE_YIELDED) = True)) Then
            If TargetCanBeReached(m_objTarget, longRange) Then
                m_Action = ACT_YIELDMONSTER
            Else
                locDebug "clsMacro.FireMageAttack: wants to Yield, but target out of range"
                Call StopCombat("FireMageAttack")
                bRet = False
                GoTo Fin
            End If
    
    'Check if we need to imperil
    ElseIf (g_ui.Macro.chkImperil.Checked _
        And (m_objTarget.UserData(B_CAN_BE_VULNED) = True) _
        And (Not m_objTarget.Imperiled)) Then
            If TargetCanBeReached(m_objTarget, longRange) Then
                m_Action = ACT_IMPERIL
            Else
                locDebug "clsMacro.FireMageAttack: wants to Imp, but target out of range"
                Call StopCombat("FireMageAttack")
                bRet = False
                GoTo Fin
            End If
            
    'Check if we need to vuln
    'If the monster hasn't been vulned to anything yet...
    ElseIf (g_ui.Macro.chkVuln.Checked _
        And (m_objTarget.CountVulns <= 0) _
        And (m_objTarget.UserData(B_CAN_BE_VULNED) = True)) Then
            'if we did find a valid vuln spell
            If g_Spells.SetupVulnForDamage(m_objTarget.Vulnerability) Then
                If TargetCanBeReached(m_objTarget, longRange) Then
                    m_Action = ACT_VULN
                Else
                    locDebug "clsMacro.FireMageAttack: wants to Vuln, but target out of range"
                    Call StopCombat("FireMageAttack")
                    bRet = False
                    GoTo Fin
                End If
            Else 'couldn't find valid vuln spell!
                PrintMessage "Could not find a valid Vuln spell for " & m_objTarget.Name
                If TargetCanBeReached(m_objTarget) Then
                    m_Action = ACT_WAR_PRIMARY
                    Call g_Spells.SetupWarSpellForDamage(m_objTarget.Vulnerability, m_objTarget.GetSquareRange)
                Else
                    locDebug "FireMageAttack : Target can not be reached with War spell"
                    Call StopCombat("FireMageAttack")
                    bRet = False
                    GoTo Fin
                End If
            End If
    
    'once all the yielding/imperiling/vulning is done...
    Else
    
        'if we're on Debuff All First mode, check to see if there are other mobs to vuln nearby
        'if not, shoot currenttarget
        If g_ui.Macro.chkDebuffFirst.Checked Then
            If FindNonDebuffedTarget(, , "FireMageAttack") Then
                locDebug "FireMageAttack: chkDebuffFirst: monster vulned, switching to another one."
                Call StopCombat("FireMageAttack - chkDebuffFirst")
                GoTo Fin
            ElseIf g_ui.Macro.chkNoWar.Checked Then
                locDebug "FireMageAttack: chkNoWar checked: stopping current combat."
                Call m_objTarget.SetUserData(B_WAR_CHECK, True)
                Call StopCombat("FireMageAttack - chkNoWar")
                m_Action = ACT_NONE
                bRet = False
                GoTo Fin
            End If
        'if we're setup to not shoot bolts at monsters...
        ElseIf g_ui.Macro.chkNoWar.Checked Then
            locDebug "FireMageAttack: chkNoWar checked: stopping current combat."
            Call m_objTarget.SetUserData(B_WAR_CHECK, True)
            Call StopCombat("FireMageAttack - chkNoWar")
            m_Action = ACT_NONE
            bRet = False
            GoTo Fin
        End If
            
        'Choose a war spell
        Dim fRange As Single
        
        fRange = m_objTarget.GetSquareRange
        m_Action = ACT_WAR_PRIMARY
        
        If TargetCanBeReached(m_objTarget) Then
            'ok, keep going
        Else
            MyDebug "FireMageAttack : Target can not be reached with War spell"
            Call StopCombat("FireMageAttack")
            bRet = False
            GoTo Fin
        End If
        
        'Check m_lastSpellTarget to see if it's the same, so we can skip all these extra steps
        
        If (m_objTarget.CountVulns > 0) Then
            If Not m_objTarget.IsVulnedTo(m_objTarget.Vulnerability) Then
                'if the monster is vulned to another vuln than the one we're using, switch our
                'bolt damage type to match the vuln if available.
                For i = DMG_SLASHING To DMG_LIGHTNING
                    If m_objTarget.IsVulnedTo(i) Then
                        If g_Spells.SetupWarSpellForDamage(i, fRange, True, False) Then
                            locDebug "FireMageAttack: " & m_objTarget.Name & " is vulned to " & GetDamageString(i)
                            Exit For
                        End If
                    End If
                Next i
            Else
                'Critter is Vulnerable to right damage type
                Call g_Spells.SetupWarSpellForDamage(m_objTarget.Vulnerability, fRange, True, False)
            End If
        Else
            'Use Vuln from UI
            Call g_Spells.SetupWarSpellForDamage(m_objTarget.Vulnerability, fRange, True, True)
        End If
        
    End If

    'Do the casting
    Select Case m_Action
        Case ACT_WAR_PRIMARY
            locDebug "clsMacroCombat.FireMageAttack: ACT_WAR_PRIMARY: " & g_Spells.CurrentWarSpell.SpellName
            Call g_Spells.CastThisSpell(g_Spells.CurrentWarSpell, m_objTarget.Guid)
        Case ACT_VULN
            locDebug "clsMacroCombat.FireMageAttack: ACT_VULN: " & g_Spells.CurrentVuln.SpellName
            Call g_Spells.CastThisSpell(g_Spells.CurrentVuln, m_objTarget.Guid)
        Case ACT_YIELDMONSTER
            locDebug "clsMacroCombat.FireMageAttack: ACT_YIELD CastCreatureDebuff"
            Call g_Spells.CastCreatureDebuff("Magic Yield Other", g_Data.mMaxVulnLevel, m_objTarget.Guid)
        Case ACT_IMPERIL
            locDebug "clsMacroCombat.FireMageAttack: ACT_IMPERIL CastLifeDebuff"
            Call g_Spells.CastLifeDebuff("Imperil Other", g_Data.mMaxVulnLevel, m_objTarget.Guid)
        Case ACT_NONE
            locDebug "clsMacroCombat.FireMageAttack: m_Action set to NONE"
    End Select
    
Fin:
    FireMageAttack = bRet
    locDebug "clsMacroCombat.FireMageAttack done, return: " & bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage "clsMacroCombat.FireMageAttack - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

'**************************************
' MeleeCheckSelection
'
' While in melee/archer macro mode,
' we don't want to loose the target selection box
' because we need to have something selected when pushing the attack key
'
' Makes sure the current target is selected in the AC client
'**************************************
Private Sub MeleeCheckSelection()
    
    If Not Valid(m_objTarget) Then
        locDebug "WARNING - MeleeCheckSelection : no valid m_objTarget!"
        Call StopCombat("MeleeCheckSelection invalid m_objTarget")
        Exit Sub
    End If

    If g_Hooks.CurrentSelection <> m_objTarget.Guid Then
        locDebug "Target is not selected - Forcing selection."
        Call g_Service.SelectObject(m_objTarget)
    End If
    
End Sub

'**************************************
' EngagingBetterTarget
'
' Looks up for a potentialy better target around the macro
' and engage it if found
'
' Returns true if the new target has been engaged
'**************************************
Private Function EngagingBetterTarget() As Boolean
On Error GoTo Error_Handler
    
    Dim bRes As Boolean
    Dim objNewTarget As acObject
    
    If Not Valid(m_objTarget) Then
        EngagingBetterTarget = False
        Exit Function
    End If
    
    If g_ui.Macro.chkAttackSelected.Checked Then
        If SelectedTargetAvailable(objNewTarget) Then
            bRes = EngageCombat(objNewTarget, "SelectedTarget")
        Else
            bRes = False
        End If
    ElseIf g_ui.Macro.chkSwitchIfBetterTarget.Checked Then
        If BetterTargetAvailable(m_objTarget, objNewTarget) Then
            If Not g_ui.Options.chkFilterLTMsg.Checked Then
                PrintMessage "Better target available - Switching from " & m_objTarget.Name & " to " & objNewTarget.Name & "..."
            Else
                MyDebug "clsMacroCombat.EngageBetterTarget: Switching from " & m_objTarget.Name & " to " & objNewTarget.Name & "..."
            End If
            bRes = EngageCombat(objNewTarget, "EngageBetterTarget")
        Else
            bRes = False
        End If
    Else
        bRes = False
    End If

Fin:
    Set objNewTarget = Nothing
    EngagingBetterTarget = bRes
    Exit Function
Error_Handler:
    bRes = False
    PrintErrorMessage "clsMacroCombat.EngageBetterTarget - " & Err.Description
    Resume Fin
End Function

'**************************************
' EngagingDangerousTarget
'
' Looks up for a target within danger zone
' and engage it if found
'
' Returns true if the new target has been engaged
'**************************************
Private Function EngagingDangerousTarget() As Boolean
On Error GoTo Error_Handler
    
    Dim bRes As Boolean
    Dim objNewTarget As acObject
    
    'default
    bRes = False
    
    If g_ui.Macro.chkDangerZone.Checked And Valid(m_objTarget) And (Not tmrDanger.Expired) Then

        Dim fRange As Double
        fRange = m_objTarget.GetSquareRange

        'If our current target is outside of the danger zone, try to find a closer target hitting us
        If fRange > (g_Data.DangerZoneRadius * g_Data.DangerZoneRadius) Then
            If FindBestTarget(objNewTarget, g_Data.DangerZoneRadius) Then
                If Not g_ui.Options.chkFilterLTMsg.Checked Then
                    PrintMessage "Danger Zone - Switching to " & objNewTarget.Name
                End If
                locDebug "clsMacroCombat.EngageDangerousTarget: Switching to " & objNewTarget.Name
                bRes = EngageCombat(objNewTarget, "EngagingDangerousTarget")
            End If
        End If

    End If
    
Fin:
    Set objNewTarget = Nothing
    EngagingDangerousTarget = bRes
    Exit Function
Error_Handler:
    bRes = False
    PrintErrorMessage "clsMacroCombat.EngagingDangerousTarget - " & Err.Description
    Resume Fin
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Local Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function GetStateString() As String
Dim sRes As String

    Select Case m_State
            
        Case ST_COMBAT_NEXTTHINK
            sRes = "Next Think"
            
        Case ST_COMBAT_ATTACKING
            sRes = "Attacking"
            
        Case ST_COMBAT_END
            sRes = "End Combat"
      
        Case Else
            sRes = "Unknown"
    End Select
    
    GetStateString = sRes
End Function

'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsMacroCombat : " & GetStateString & "] " & DebugMsg, bSilent)
    End If
End Sub

