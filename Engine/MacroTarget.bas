Attribute VB_Name = "MacroTarget"
Option Explicit

Private Const DEBUG_ME = False

'Friendly players = same allegiance or people in fellow
Public Function IsFriendlyPlayer(objPlayer As acObject)
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim aName As String
    
    If Valid(objPlayer) Then
        aName = objPlayer.Name
        If g_ui.Options.NameInFriendsList(aName) Then
            bRet = True
        Else
            If g_ui.Options.chkOnlyFriendsList.Checked Then
                bRet = False
            Else
                bRet = (objPlayer.MonarchID = g_Objects.Player.MonarchID) _
                        Or g_ds.GameObjects.Fellowship.Exists(objPlayer.Guid)
            End If
        End If
    End If

Fin:
    IsFriendlyPlayer = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "IsFriendlyPlayer - " & Err.Description
    Resume Fin
End Function

Public Function IsValidPlayerTarget(Target As acObject, Optional ByVal bProtectFriendly As Boolean = True)
On Error GoTo ErrorHandler
    
    'Default
    IsValidPlayerTarget = False
    
    If g_ui.Macro.chkAttackPK.Checked Or Valid(Target) Then
        Exit Function
    
    'Leave if we're not even pk
    ElseIf g_Objects.Player.PlayerType = PLAYER_WHITE Then
        Exit Function
        
    'Bad player target?
    ElseIf (Not IsPlayer(Target)) Or g_Objects.IsSelf(Target) Then
        Exit Function
        
    'Not same PK type?
    ElseIf Target.PlayerType <> g_Objects.Player.PlayerType Then
        Exit Function
    
    'Friendly player, don't shoot him !
    ElseIf bProtectFriendly And IsFriendlyPlayer(Target) Then
        Exit Function
    
    'Alright, let's mess !
    Else
        IsValidPlayerTarget = True
        
    End If

Fin:
    Exit Function
ErrorHandler:
    IsValidPlayerTarget = False
    PrintErrorMessage "IsValidPlayerTarget - " & Err.Description
    Resume Fin
End Function

Public Function IsValidTarget(ByVal Target As acObject, Optional bMustBeInList As Boolean = False, Optional bIgnoreBlacklist As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    
    'default ret val
    IsValidTarget = False
    
    'locDebug "IsValidTarget(" & Target.Name & ") : start"
    If Not Valid(Target) Then
        locDebug "IsValidTarget() : target is nothing"
        Exit Function
        
    'Already dead ?
    ElseIf Target.Dead Then
        Exit Function
        
    'Not a monster or valid player?
    ElseIf Not (IsMonster(Target) Or IsValidPlayerTarget(Target)) Then
        locDebug "IsValidTarget() : " & Target.Name & " is not a monster or valid player target"
        Exit Function
    
    'Blacklisted?
    ElseIf (Not bIgnoreBlacklist) And (Target.UserData(B_DANGEROUS) = False) And (Target.UserData(INT_BLIST_TIME) > g_Core.Time) Then
        MyDebug "IsValidTarget(" & Target.Name & ") : target is blacklisted"
        Exit Function
        
    'Valid monster ?
    ElseIf bMustBeInList And IsMonster(Target) And (Target.UserData(B_DANGEROUS) = False) And (Not MonsterEnabled(Target)) Then
        locDebug "IsValidTarget(" & Target.Name & ") : monster not enabled"
        Exit Function
        
    ElseIf g_ui.Macro.chkNoWar.Checked And (Target.UserData(B_WAR_CHECK) = True) Then
        locDebug "IsValidTarget(" & Target.Name & ") : No War is checked"
        Exit Function
        
    'Ok - let's shoot it !
    Else
        IsValidTarget = True
    End If
    
Fin:
    Exit Function
ErrorHandler:
    IsValidTarget = False
    PrintErrorMessage "IsValidTarget - " & Err.Description
    Resume Fin
End Function

'SelectedTargetAvailable
Public Function SelectedTargetAvailable(ByRef objTargetOut As acObject) As Boolean
On Error GoTo ErrorHandler
    
    'Default to false
    SelectedTargetAvailable = False
    
    If g_Hooks.CurrentSelection = 0 Then
        Exit Function
    End If
        
    Set objTargetOut = g_Objects.FindObject(g_Hooks.CurrentSelection)
    
    If IsMonster(objTargetOut) Then
        SelectedTargetAvailable = True
    Else
        locDebug "SelectedTargetAvailable - invalid objNewTarget"
        Exit Function
    End If

Fin:
    Exit Function
ErrorHandler:
    SelectedTargetAvailable = False
    PrintErrorMessage "SelectedTargetAvailable - " & Err.Description
    Resume Fin
End Function


Public Function BetterTargetAvailable(ByVal objCurTarget As acObject, ByRef objTargetOut As acObject) As Boolean
On Error GoTo ErrorHandler
    
    Dim objNewTarget As acObject
    Dim bFoundTarget As Boolean
    
    'Default to false
    BetterTargetAvailable = False
    
    If Not Valid(objCurTarget) Then
        MyDebug "WARNING - BetterTargetAvailable - invalid objCurTarget"
        Exit Function
    End If
    
    'Find the best potential target around us
    bFoundTarget = FindBestTarget(objNewTarget)
    
    If bFoundTarget Then
        If Not Valid(objNewTarget) Then
            PrintErrorMessage "BetterTargetAvailable - invalid objNewTarget"
            Exit Function
        End If
        
        If objNewTarget.Guid = objCurTarget.Guid Then
            'well... we're already on it :)
            Exit Function
        End If
        
        'Check if it's worth switching to it
        If objNewTarget.Priority > objCurTarget.Priority Then
            'yes definately worth it
            Set objTargetOut = objNewTarget
            BetterTargetAvailable = True
            locDebug "MacroTarget.BetterTargetAvailable: Higher Priority: " & objNewTarget.Name
        ElseIf objNewTarget.Priority = objCurTarget.Priority Then
            'if same priority, only switch if our current target isnt vulned and the other is
            If (objCurTarget.Vulns <= 0) And (objNewTarget.Vulns > 1) Then
                Set objTargetOut = objNewTarget
                BetterTargetAvailable = True
                locDebug "MacroTarget.BetterTargetAvailable: is Vulned: " & objNewTarget.Name
            End If
        Else
            Set objTargetOut = Nothing
            BetterTargetAvailable = False
        End If
    End If
    
Fin:
    Set objNewTarget = Nothing
    Exit Function
ErrorHandler:
    BetterTargetAvailable = False
    PrintErrorMessage "BetterTargetAvailable - " & Err.Description
    Resume Fin
End Function

'Target scanner function that looks for a Valid target within Range
Public Function TargetScanner(ByVal colTargets As colObjects) As Boolean
On Error GoTo ErrorHandler
    Dim fSearchRadius As Integer
    Dim objEntity As acObject

    If g_Macro.CombatType = TYPE_ARCHER Then
            fSearchRadius = g_ui.Macro.txtArcherRadius.Text
    ElseIf g_Macro.CombatType = TYPE_MELEE Then
            fSearchRadius = g_ui.Macro.txtMeleeRadius.Text
    Else
        fSearchRadius = g_ui.Macro.txtMageRadius.Text
        If g_ui.Macro.chkVuln.Checked And fSearchRadius < g_ui.Macro.txtVulnRange.Text Then
            fSearchRadius = g_ui.Macro.txtVulnRange.Text
        End If
    End If


    'If g_ui.Macro.chkDebuffFirst.Checked And IsCaster Then
    '    bFound = FindNonDebuffedTarget(objTargetOut, fSearchRadius, "FindBestTarget")
    '    If bFound Then GoTo Fin
    'End If
    '
    'bFound = FindTarget(objTargetOut, fSearchRadius, Not g_ui.Macro.chkAttackAny.Checked, , AttackVulnedMobsFirst, , "FindBestTarget")
    
    'loop through each object of the collection
    For Each objEntity In colTargets
        'First make sure it's a valid potential target
        If IsValidTarget(objEntity, True) Then
            'Is Target within range? -
            If TargetCanBeReached(objEntity, fSearchRadius) Then
                'locDebug "TargetScanner: " & objEntity.Name & " is in range -- SR:" & fSearchRadius
                TargetScanner = True
                GoTo Fin
            End If
        End If
    Next objEntity
    
    TargetScanner = False

Fin:
    Set objEntity = Nothing
    Exit Function
ErrorHandler:
    TargetScanner = False
    PrintErrorMessage "TargetScanner - " & Err.Description
    Resume Fin
End Function


Public Function FindBestTarget(Optional ByRef objTargetOut As acObject, Optional ByVal fSearchRadius As Single = -1) As Boolean
On Error GoTo ErrorHandler

    Dim bFound As Boolean

    If g_ui.Macro.chkDebuffFirst.Checked And IsCaster Then
        bFound = FindNonDebuffedTarget(objTargetOut, fSearchRadius, "FindBestTarget")
        If bFound Then GoTo Fin
    End If
    
    bFound = FindTarget(objTargetOut, fSearchRadius, Not g_ui.Macro.chkAttackAny.Checked, , AttackVulnedMobsFirst, , "FindBestTarget")
    
Fin:
    FindBestTarget = bFound
    Exit Function
ErrorHandler:
    bFound = False
    PrintErrorMessage "FindBestTarget - " & Err.Description
    Resume Fin
End Function

Public Function FindRingSpellTarget(Optional ByVal fSearchRadius As Single = 10) As Boolean
On Error GoTo ErrorHandler

    Dim bFound As Boolean
    Dim iCount As Integer
    Dim colTargets As colObjects
    Dim objEntity As acObject
    
    iCount = 0
    
    'get list of all critters around us
    Set colTargets = g_Objects.Monsters
    
    'loop through each Monster and check range
    For Each objEntity In colTargets
        If Valid(objEntity) And TargetCanBeReached(objEntity, fSearchRadius) Then
            iCount = iCount + 1
        End If
    Next objEntity
    
    If (iCount > g_ui.Macro.txtRingNum.Text) Then
        bFound = True
    End If
    
    'MyDebug "FindRingSpellTarget: found monsters: " & iCount

Fin:
    FindRingSpellTarget = bFound
    Set objEntity = Nothing
    Set colTargets = Nothing
    Exit Function
ErrorHandler:
    bFound = False
    PrintErrorMessage "FindRingSpellTarget - " & Err.Description
    Resume Fin
End Function

Public Function TargetCanBeReached(ByVal objTarget As acObject, Optional ByVal AttackRange As Single = -1, Optional ByRef fSquareRangeOut As Single) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    Dim TargetZoff, PlayerZoff As Long
    
    If Not Valid(objTarget) Then
        PrintErrorMessage "TargetCanBeReached - invalid objTarget"
        GoTo Fin
    End If

    If AttackRange = -1 Then
        If g_Macro.CombatType = TYPE_ARCHER Then
            AttackRange = g_ui.Macro.txtArcherRadius.Text
        ElseIf g_Macro.CombatType = TYPE_MELEE Then
            AttackRange = g_ui.Macro.txtMeleeRadius.Text
        Else
            AttackRange = g_ui.Macro.txtMageRadius.Text
        End If
    End If
     
    TargetZoff = Round(objTarget.Loc.Zoff, 1)
    PlayerZoff = Round(g_Objects.Player.Loc.Zoff, 1)

    If ValidRangeTo(objTarget, AttackRange, fSquareRangeOut) Then
        If ((g_ui.Macro.chkAttackGroundOnly.Checked) And (Abs(TargetZoff - PlayerZoff) > 1)) Then
            locDebug "Target not on same ground level, can't be reached."
            bRet = False
        Else
            locDebug "TargetCanBeReached:YES: " & AttackRange & "::" & fSquareRangeOut & " : " & objTarget.Name & " : " & objTarget.Guid
            bRet = True
        End If
    Else
        locDebug "TargetCanBeReached:NO: " & AttackRange & "::" & fSquareRangeOut & " : " & objTarget.Name & " : " & objTarget.Guid
        bRet = False
    End If
    
Fin:
    TargetCanBeReached = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "TargetCanBeReached - " & Err.Description
    Resume Fin
End Function

Public Function FindTargetInCol(colTargets As colObjects, _
                                Optional ByRef objTargetOut As acObject, _
                                Optional ByVal SearchRadius As Single = -1, _
                                Optional ByVal bMustBeInList As Boolean = False, _
                                Optional ByVal bDanger As Boolean = False, _
                                Optional ByVal bPriorityToVulneds As Boolean = False, _
                                Optional ByVal bNonDebuffedTargetsOnly As Boolean = False, _
                                Optional ByVal sSource As String = "") As Boolean
On Error GoTo ErrorHandler

    Dim bFoundTarget As Boolean
    Dim objEntity As acObject
    
    'NOTE:
    'Ranges are actually Square Ranges, but it doesnt matter here since we're only interested in
    'comparing two ranges values
    
    Dim fRange As Single
    Dim fBestTargetRange As Single      'range of the currently selected target
    Dim iCurBestPriority As Integer     'priority of the currently selected target
    Dim bCurVulned As Boolean           'is current object vulned?
    Dim bPrevVulned As Boolean          'is previous best target vulned?
    Dim objBestTarget As acObject       'the current best target we found
    Dim bDontCompareToBest As Boolean   'set to true in the case where the current best target is not vulned,
                                        'and the current target is vulned - used to ignore priority comparisons
                                        'since we want to pick the vulned target even if it has lower priority
    
    'Init
    Set objBestTarget = Nothing
    iCurBestPriority = 0
    fBestTargetRange = 99999
    bCurVulned = False
    bPrevVulned = False
    bFoundTarget = False
    
    'loop through each object of the collection
    For Each objEntity In colTargets
           
        'First make sure it's a valid potential target
        If Not IsValidTarget(objEntity, bMustBeInList, bDanger) Then
            locDebug "FindTargetInCol - [ " & objEntity.Name & " ] -  NOT Valid Target"
            GoTo NextEntity
        
        'If we're in "Debuff All Monsters First" and this target is already debuffed, switch to next one
        ElseIf bNonDebuffedTargetsOnly And TargetDebuffed(objEntity) Then
            locDebug "FindTargetInCol - [ " & objEntity.Name & " ] - bNonDebuffedTargetsOnly And TargetDebuffed(objEntity)"
            GoTo NextEntity
            
        Else
    
            'Reset bCurVulned
            bCurVulned = False
            bDontCompareToBest = False
            
            'If this monster is considered vulned
            If bPriorityToVulneds And ((objEntity.Vulns > 0) Or (IsMelee And objEntity.Imperiled)) Then
                locDebug "FindTargetInCol: AttackVulnedMobsFirst: " & objEntity.Name & " is vulned."
                bCurVulned = True
                
                'If it's the first vulned monster we find and we give priority to vulned targets,
                'set the bDontCompareToBest flag so that we won't compare the priority of this target
                'against the priority of the current best target (which is not vulned)
                If (Not bPrevVulned) And Valid(objBestTarget) Then
                    bDontCompareToBest = True
                End If
                    
            End If
            
            locDebug "FindTargetInCol(" & sSource & ") - Testing [ " & objEntity.Name & " ] P:" & objEntity.Priority & " - Vulned: " & bCurVulned & " - First Vulned? " & bDontCompareToBest
            
            'If we already have a potential target, check to see if we can skip this one before computing range
            If Valid(objBestTarget) Then
            
                locDebug "...Current Best Target is " & objBestTarget.Name & " [P:" & objBestTarget.Priority & "] @ " & fBestTargetRange & " - Vulned: " & bPrevVulned
                
                'If we're only interested in vulned targets and this one is not vulned, skip it
                If bPrevVulned And (Not bCurVulned) Then
                    'locDebug "........Skipping: bPrevVulned And (Not bCurVulned)"
                    GoTo NextEntity
                    
                'If we're lower priority than the currently best target, skip
                ElseIf (Not bDontCompareToBest) And objEntity.Priority < iCurBestPriority Then
                    'locDebug "........Skipping: (Not bDontCompareToBest) And objEntity.Priority < iCurBestPriority"
                    GoTo NextEntity
                    
                End If
                
            End If
                    
            'Is Target within range? -
            'Also put the square range to objEntity in fRange so that we don't have to recompute it further
            If Not TargetCanBeReached(objEntity, SearchRadius, fRange) Then
                'locDebug "........Skipping: " & objEntity.Name & " can NOT be reached -- SR:" & SearchRadius & "  fRange:" & fRange
                GoTo NextEntity
                
            Else
    
                'If we have the same priority than the currently best target, but are further away, skip
                If (Not bDontCompareToBest) And (objEntity.Priority = iCurBestPriority) And (fRange > fBestTargetRange) Then
                    'locDebug "........Skipping " & objEntity.Name & " [P:" & objEntity.Priority & "] @ " & fRange & " - Best Target Closer @ " & fBestTargetRange
                    GoTo NextEntity
                    
                'Else we're the new best target
                Else
                
                    Set objBestTarget = objEntity
                    fBestTargetRange = fRange
                    iCurBestPriority = objEntity.Priority
                    bPrevVulned = bCurVulned
                    bFoundTarget = True
                    locDebug ">>>>>>>> Setting As Best Target [Range: " & fBestTargetRange & " - P: " & iCurBestPriority & " - Vulned: " & bCurVulned & "]"
                    
                End If
                
            End If
            
        End If
        
NextEntity:
    Next objEntity
    
    If bFoundTarget Then
       'If g_ui.Macro.chkNoWar.Checked Then Call objBestTarget.SetUserData(B_WAR_CHECK, True)
       locDebug "FindTargetInCol (" & sSource & "): selecting " & objBestTarget.Name & " [" & objBestTarget.Guid & "] at Range: " & fBestTargetRange
       If bDanger Then Call objBestTarget.SetUserData(B_DANGEROUS, True)
    End If

Fin:
    Set objTargetOut = objBestTarget
    FindTargetInCol = bFoundTarget
    Set objEntity = Nothing
    Set objBestTarget = Nothing
    Exit Function
ErrorHandler:
    bFoundTarget = False
    PrintErrorMessage "FindTargetInCol - " & Err.Description
    Resume Fin
End Function

Public Function FindTarget(Optional ByRef objTargetOut As acObject, _
                                Optional ByVal SearchRadius As Single = -1, _
                                Optional ByVal bMustBeInList As Boolean = False, _
                                Optional ByVal bDanger As Boolean = False, _
                                Optional ByVal bPriorityToVulneds As Boolean = False, _
                                Optional ByVal bNonDebuffedTargetsOnly As Boolean = False, _
                                Optional ByVal sSource As String = "") As Boolean
On Error GoTo ErrorHandler

    Dim bFoundTarget As Boolean
    
    If g_ui.Macro.chkAttackPK.Checked Then
        'look for PKs before looking for monsters
        bFoundTarget = FindTargetInCol(g_Objects.Players, objTargetOut, SearchRadius, bMustBeInList, bDanger, bPriorityToVulneds, bNonDebuffedTargetsOnly, sSource)
    End If
    
    'look for monster targets
    If Not bFoundTarget Then bFoundTarget = FindTargetInCol(g_Objects.Monsters, objTargetOut, SearchRadius, bMustBeInList, bDanger, bPriorityToVulneds, bNonDebuffedTargetsOnly, sSource)

Fin:
    FindTarget = bFoundTarget
    Exit Function
ErrorHandler:
    bFoundTarget = False
    PrintErrorMessage "FindTarget - " & Err.Description
    Resume Fin
End Function

'Returns true if all the debuffs allowed are active on this target
Public Function TargetDebuffed(ByVal Target As acObject, Optional bVuln As Boolean = True, Optional bImp As Boolean = True) As Boolean
On Error GoTo ErrorHandler

    Dim bRes As Boolean
    
    If Not Valid(Target) Then
        PrintErrorMessage "TargetDebuffed - invalid Target"
        GoTo Fin
    End If

    bRes = True
    
    If bRes And g_ui.Macro.chkImperil.Checked And (Target.UserData(B_CAN_BE_IMPERILED) = True) Then
        bRes = Target.Imperiled
    End If
    
    If bRes And g_ui.Macro.chkVuln.Checked And (Target.UserData(B_CAN_BE_VULNED) = True) Then
        bRes = (Target.Vulns > 0)
    End If
    
    If bRes And g_ui.Macro.chkYield.Checked And (Target.UserData(B_CAN_BE_YIELDED) = True) Then
        bRes = Target.Yielded
    End If

Fin:
    TargetDebuffed = bRes
    Exit Function
ErrorHandler:
    bRes = False
    PrintErrorMessage "TargetDebuffed - " & Err.Description
    Resume Fin
End Function

Public Function FindNonDebuffedTarget(Optional ByRef objTargetOut As acObject, Optional ByVal fSearchRadius As Single = -1, Optional sSource As String = "") As Boolean
On Error GoTo ErrorHandler
    
    'MyDebug "FindNonDebuffedTarget: fSearchRadius: " & fSearchRadius
    FindNonDebuffedTarget = FindTarget(objTargetOut, fSearchRadius, Not g_ui.Macro.chkAttackAny.Checked, , False, True, "[FindNonDebuffedTarget] " & sSource)

Fin:
    Exit Function
ErrorHandler:
    FindNonDebuffedTarget = False
    PrintErrorMessage "FindNonDebuffedTarget - " & Err.Description
    Resume Fin
End Function

'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[MacroTarget] " & DebugMsg, bSilent)
    End If
End Sub
