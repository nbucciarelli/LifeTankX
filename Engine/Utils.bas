Attribute VB_Name = "Utils"
Option Explicit


Private Const DEBUG_ME = False


Public Sub MyDebug(strMsg As String, Optional bSilent As Boolean = False)
On Error GoTo ErrorHandler

    Call g_Engine.FireDebugMessage(strMsg)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ERROR in MyDebug[" & strMsg & "] - " & Err.Description
    Resume Fin
End Sub

Public Sub LogEvent(strMsg As String, Optional bSilent As Boolean = False)
On Error GoTo ErrorHandler

    Call g_Engine.FireLogEvent(strMsg)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "LogEvent - " & Err.Description
    Resume Fin
End Sub

Public Sub PrintMessage(strMsg As String, Optional Color As Long = COLOR_CYAN)
On Error GoTo ErrorHandler

    Call g_Engine.FirePrintMessage(strMsg, Color)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ERROR in PrintMessage[" & strMsg & "] - " & Err.Description
    Resume Fin
End Sub

Public Sub PrintWarning(strMsg As String)
On Error GoTo ErrorHandler

   Call g_Engine.FireWarningMessage(strMsg)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PrintWarning(" & strMsg & ") - " & Err.Description
    Resume Fin
End Sub

Public Sub PrintErrorMessage(ByVal strMsg As String, Optional ByVal bShowErrorNum As Boolean = True)
On Error GoTo ErrorHandler
    
    Call g_Engine.FireErrorMessage(strMsg)
    
Fin:
    Exit Sub
ErrorHandler:
    MsgBox "(Engine) ERROR @ PrintErrorMessage - " & Err.Description & " - line: " & Erl & " [msg: " & strMsg & "]"
    Resume Fin
End Sub

Public Sub LogChatMessage(ByVal sMsg As String)
On Error GoTo ErrorHandler
    
    Call g_Engine.FireLogChatMessage(sMsg)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "LogChatMessage - " & Err.Description
    Resume Fin
End Sub

Public Function BoolToInteger(BoolValue As Boolean) As String
    If BoolValue = True Then
        BoolToInteger = 1
    Else
        BoolToInteger = 0
    End If
End Function

Public Function GetPercent(ByVal Source As Long, ByVal Percentage As Integer) As Long
    GetPercent = ((CLng(Percentage) * Source) / 100)
End Function

Public Function GetSquareRange(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single) As Single
    GetSquareRange = ((x2 - x1) * (x2 - x1)) + ((y2 - y1) * (y2 - y1)) + ((z2 - z1) * (z2 - z1))
End Function

Public Function GetRange(ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single) As Single
    GetRange = Sqr(GetSquareRange(x1, y1, z1, x2, y2, z2))
End Function

Public Function WorldRange(ByVal aGuid As Long) As Single
    WorldRange = (g_Filters.g_worldFilter.Distance2D(aGuid, g_Filters.PlayerGUID) * 200)
End Function

Public Function GetIntegerMinutesFromSeconds(TimeInSeconds As Long) As Integer
Dim SecondsLeft As Long
Dim NumMinutes As Integer
    
    NumMinutes = 0
    SecondsLeft = TimeInSeconds - 60
    While (SecondsLeft > 0)
        NumMinutes = NumMinutes + 1
        SecondsLeft = SecondsLeft - 60
    Wend
    
    GetIntegerMinutesFromSeconds = NumMinutes
    
End Function

'0 based index
Public Sub BuildArgsList(ByVal Data As String, ByRef Args() As String, ByRef NumArgs As Integer, Optional TokenChar As String = " ")
Dim Pos As Integer
Dim DataCopy As String
Dim CurArg As String 'current argument

    NumArgs = 0
    
    Data = Trim(Data)
    DataCopy = Data
    
    If Len(DataCopy) > 0 Then
        Do
            DataCopy = Trim(DataCopy)
            Pos = InStr(1, DataCopy, TokenChar)
            If Pos > 0 Then
                CurArg = Mid(DataCopy, 1, Pos - 1)
                ReDim Preserve Args(NumArgs + 1)
                Args(NumArgs) = CurArg
                NumArgs = NumArgs + 1
                'remove this argument from the data list, so we can extract the remaining ones
                DataCopy = Mid(DataCopy, Pos + Len(TokenChar))
            Else 'Last arg
                CurArg = Mid(DataCopy, 1)
                ReDim Preserve Args(NumArgs + 1)
                Args(NumArgs) = CurArg
                NumArgs = NumArgs + 1
            End If
        Loop While (Pos > 0 And DataCopy <> "" And CurArg <> "")
    End If
    
End Sub

Public Sub BuildPartialArgsList(ByVal Data As String, ByRef Args() As String, ByRef NumArgs As Integer, ByVal MaxArgs, Optional TokenChar As String = " ")
Dim Pos As Integer
Dim DataCopy As String
Dim CurArg As String 'current argument

    NumArgs = 0
    
    Data = Trim(Data)
    DataCopy = Data
    
    If Len(DataCopy) > 0 Then
        Do
            DataCopy = Trim(DataCopy)
            Pos = InStr(1, DataCopy, TokenChar)
            If Pos > 0 Then
                ReDim Preserve Args(NumArgs + 1)
                If (NumArgs + 1 >= MaxArgs) Then
                    CurArg = Mid(DataCopy, 1)
                    Args(NumArgs) = CurArg
                    NumArgs = NumArgs + 1
                    Exit Sub
                Else
                    CurArg = Mid(DataCopy, 1, Pos - 1)
                    Args(NumArgs) = CurArg
                    NumArgs = NumArgs + 1
                    'remove this argument from the data list, so we can extract the remaining ones
                    DataCopy = Mid(DataCopy, Pos + Len(TokenChar))
                End If
            Else 'Last arg
                CurArg = Mid(DataCopy, 1)
                ReDim Preserve Args(NumArgs + 1)
                Args(NumArgs) = CurArg
                NumArgs = NumArgs + 1
            End If
        Loop While (Pos > 0 And DataCopy <> "" And CurArg <> "")
    End If
    
End Sub

Public Function dmg2vuln(ByVal DamageType As eDamageType) As Integer
    Select Case DamageType
        Case DMG_SLASHING
            dmg2vuln = FL_SLASHING
        Case DMG_PIERCING
            dmg2vuln = FL_PIERCING
        Case DMG_BLUDGEONING
            dmg2vuln = FL_BLUDGEONING
        Case DMG_FIRE
            dmg2vuln = FL_FIRE
        Case DMG_COLD
            dmg2vuln = FL_COLD
        Case DMG_ACID
            dmg2vuln = FL_ACID
        Case DMG_LIGHTNING
            dmg2vuln = FL_LIGHTNING
        Case Else
            PrintErrorMessage "dmg2vuln: unknown damage type " & DamageType
            dmg2vuln = -1
    End Select
End Function

Public Function SearchCSVTokenPos(ByVal StartPos As Integer, ByVal Source As String) As Integer
Dim Pos As Integer
    
    Pos = InStr(StartPos, Source, ",")
    If Pos = 0 Then 'try with ;
        Pos = InStr(StartPos, Source, ";")
    End If
    
    SearchCSVTokenPos = Pos
End Function

Public Function GetMeleeMasteryName(ByVal iMasterySkill As Integer) As String
    Select Case iMasterySkill
        Case SKILL_AXE
            GetMeleeMasteryName = "Axe"
        Case SKILL_DAGGER
            GetMeleeMasteryName = "Dagger"
        Case SKILL_MACE
            GetMeleeMasteryName = "Mace"
        Case SKILL_SPEAR
            GetMeleeMasteryName = "Spear"
        Case SKILL_STAFF
            GetMeleeMasteryName = "Staff"
        Case SKILL_SWORD
            GetMeleeMasteryName = "Sword"
        Case SKILL_UNARMED_COMBAT
            GetMeleeMasteryName = "Unarmed Combat"
        Case Else
            GetMeleeMasteryName = "None"
    End Select
End Function

Public Function GetArcherMasteryName(ByVal iMasterySkill As Integer) As String
    Select Case iMasterySkill
        Case SKILL_BOW
            GetArcherMasteryName = "Bow"
        Case SKILL_CROSSBOW
            GetArcherMasteryName = "Crossbow"
        Case SKILL_THROWN_WEAPONS
            GetArcherMasteryName = "Thrown Weapons"
        Case Else
            GetArcherMasteryName = "None"
    End Select
End Function

Public Function SkillToSpellLevel(iSkill As Integer) As Integer
Dim iVal As Integer

    If iSkill >= 400 Then
        iVal = 8
    ElseIf iSkill >= 300 Then
        iVal = 7
    ElseIf iSkill >= 250 Then
        iVal = 6
    ElseIf iSkill >= 200 Then
        iVal = 5
    ElseIf iSkill >= 150 Then
        iVal = 4
    ElseIf iSkill >= 100 Then
        iVal = 3
    ElseIf iSkill >= 50 Then
        iVal = 2
    Else
        iVal = 1
    End If

    SkillToSpellLevel = iVal
    
End Function


Public Function Arccos(ByVal x As Single) As Single
    Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

'All vectors are 0 based
Public Function VecLen2D(ByVal x As Single, ByVal y As Single) As Single
    VecLen2D = Sqr(x * x + y * y)
End Function

Public Sub VecNormalize2D(ByRef x As Single, ByRef y As Single)
Dim fLen As Single

    fLen = VecLen2D(x, y)
    If fLen <= 0 Then
        MyDebug "VecNormalize2D: error fLen <=0"
        Exit Sub
    End If
    x = x / fLen
    y = y / fLen
    
End Sub

'Returns angle in rad
'Public Function VecToAngle2D(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
'    Dim fLen As Single
'
'    Call VecNormalize2D(x1, y1)
'    Call VecNormalize2D(x2, y2)
'
'    fLen = VecLen2D(x1, y1) * VecLen2D(x2, y2)
'
'    If fLen <= 0 Then
'        MyDebug "VecToAngle2D: Error fLen <= 0"
'        VecToAngle2D = 0
'        Exit Function
'    End If
'
'    VecToAngle2D = Arccos((x1 * x2 + y1 * y2) / fLen)
'End Function

Public Function VecToAngle2D(ByVal v1 As clsVector, ByVal v2 As clsVector) As Single
    Dim fLen As Single
    
    v1.Normalize2D
    v2.Normalize2D
    fLen = v1.Len2D * v2.Len2D
    
    If fLen <= 0 Then
        MyDebug "VecToAngle2D: Error fLen <= 0"
        VecToAngle2D = 0
        Exit Function
    End If
    
    VecToAngle2D = Arccos((v1.x * v2.x + v1.y * v2.y) / fLen)
End Function

Public Function DegToRad(ByVal fAngleInDeg As Single) As Double
    DegToRad = (M_PI * fAngleInDeg) / 180
End Function

Public Function RadToDeg(ByVal fAngleInRad As Double) As Double
    RadToDeg = (180 * fAngleInRad) / M_PI
End Function

' Normalize the given turning angle to within plus-or-minus 180 degrees.
Public Function DheadNormalize(ByVal dhead As Single) As Single
    
    Dim dheadT: dheadT = dhead
    Do While dheadT > 180
        dheadT = dheadT - 360
    Loop
    Do While dheadT < -180
        dheadT = dheadT + 360
    Loop
    DheadNormalize = dheadT
    
End Function

Public Property Get curHeading() As Single

    curHeading = g_Hooks.HeadingDegrees
    
    'curHeading = g_Hooks.Heading
    'curHeading = 180 - curHeading * 360 / (2 * M_PI)
    'If curHeading < 0 Then curHeading = curHeading + 360
    
End Property

Public Function landblockInRange(ByVal landblock As Long) As Boolean
        
        Dim ilbLng As Long: ilbLng = Int(landblock / &H1000000) And &HFF&
        Dim ilbLat As Byte: ilbLat = Int(landblock / &H10000) And &HFF&
        
        Dim isqLngMbr As Long, isqLatMbr As Long
        Dim isqLng As Long, isqLat As Long
        
        isqLngMbr = g_ds.AccuratePlayerLoc.Longitude
        isqLatMbr = g_ds.AccuratePlayerLoc.Latitude
        
        If (landblock And &HFF00&) = 0 Then
        '   Target is outdoors; extract square numbers.
            Dim dwLow As Long
            dwLow = (landblock And &HFF&) - 1
            isqLng = ilbLng * 8 + (dwLow \ 8 And &H7)
            isqLat = ilbLat * 8 + (dwLow And &H7)
        '   Do short-range check (+/- 4 squares)
            landblockInRange = _
              (Abs(isqLng - isqLngMbr) < 4) Or _
              (Abs(isqLat - isqLatMbr) < 4)
        Else
        '   Target is indoors; use central square of landblock.
            isqLng = ilbLng * 8 + 4
            isqLat = ilbLat * 8 + 4
        '   Do long-range check (+/- 9 squares)
            landblockInRange = _
              (Abs(isqLat - isqLatMbr) < 9) Or _
              (Abs(isqLng - isqLngMbr) < 9)
        End If
    
End Function


' -----------------

'Rounds down a number and returns an integer
Public Function IntRoundDown(ByVal fNum As Double) As Integer
    Dim iRounded As Integer
        
    '0.654 -> CInt(0.654) = 1 -> fDelta = 1 - 0.654 > 0 -> round down
    '0.342 -> CInt(0.342) = 0 -> fDelta = 0 - 0.342 < 0 -> dont round down
    
    iRounded = CInt(fNum)
    If (iRounded - fNum) > 0 Then
        IntRoundDown = iRounded - 1
    Else
        IntRoundDown = iRounded
    End If

End Function

Public Function ACBusy() As Boolean
   If Valid(g_Hooks) And Valid(g_Spells) Then
       ACBusy = (g_Hooks.BusyState <> 0) Or (g_Spells.Casting)
   End If
End Function

Public Function IsBusy(Optional ByVal bTurbo As Boolean = False) As Boolean
    Dim bMacroBusy As Boolean
    bMacroBusy = False
    
    'IsBusy = (g_Spells.Casting) Or (g_ds.InPortalSpace) Or (Not g_Macro.Combat.AttackMoveCompleted)
    IsBusy = (g_ds.InPortalSpace) Or (Not g_Macro.Combat.AttackMoveCompleted)
    
    If Not bTurbo Then
        If Valid(g_Macro) Then
            If g_Macro.Ticking Then
                bMacroBusy = Not g_Macro.BusyDelayTimer.Expired
            End If
        End If
        IsBusy = IsBusy Or ACBusy Or bMacroBusy
    End If
    
    'If g_Spells.Casting And Not (g_Macro.PostCastDelayTimer.Expired) Then
    '    IsBusy = True
    'End If
    
End Function

Public Function FormatXp(Val As Variant, Optional ShowXpTag As Boolean = True, Optional szSeparator As String = ",") As String
    Dim Tag As String
    Dim tmp As String
    Dim sLen As Integer
    Dim i As Integer
    
    FormatXp = FormatNumber(Val, 0, 0, 0, True)
    
    If ShowXpTag Then
        FormatXp = FormatXp & " xp"
    End If

End Function

Public Sub ChoiceListSelect(ByVal EntryText As String, ByRef choiceList As DecalControls.Choice)
Dim i As Integer
    
    For i = 0 To choiceList.ChoiceCount - 1
        If SameText(choiceList.Text(i), EntryText) Then
            choiceList.Selected = i
            Exit Sub
        End If
    Next i
    
End Sub

Public Function GetEquipPartString(ByVal iEquipPart As eEquipParts)
Dim sRet As String

    Select Case iEquipPart
        Case EQ_NONE
            sRet = "None"
        Case EQ_HEAD
            sRet = "Head"
        Case EQ_FEET
            sRet = "Feet"
        Case EQ_HANDS
            sRet = "Hands"
        Case EQ_TOP
            sRet = "Top"
        Case EQ_BOTTOM
            sRet = "Bottom"
        Case EQ_UNDERWEAR_TOP
            sRet = "Top Underwears"
        Case EQ_UNDERWEAR_BOTTOM
            sRet = "Bottom Underwears"
        Case EQ_SHIELD
            sRet = "Shield"
        Case EQ_WEAPON
            sRet = "Weapon"
        Case EQ_WAND
            sRet = "Wand"
        Case Else
            sRet = "Unknown Equip Part " & iEquipPart
    End Select
    
    GetEquipPartString = sRet
End Function

Public Function CreateTimer() As clsTimer
On Error GoTo ErrorHandler

    If Valid(g_Timers) Then
        Set CreateTimer = g_Timers.CreateTimer
    Else
        Set CreateTimer = Nothing
    End If

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "CreateTimer - " & Err.Description
    Resume Fin
End Function

Public Sub ToggleCombatMode(newCombatState As eCombatStates)
    
    If (g_Hooks.CombatMode = 1) Then
        Call g_Hooks.SetCombatMode(newCombatState)
    Else
        Call g_Hooks.SetCombatMode(COMBATSTATE_PEACE)
    End If
    
End Sub

Public Function IsMelee() As Boolean
    IsMelee = (g_Macro.CombatType = TYPE_MELEE) Or (g_Macro.CombatType = TYPE_ARCHER)
End Function

Public Function IsCaster() As Boolean
    IsCaster = (g_Macro.CombatType = TYPE_CASTER)
End Function

Public Sub ToggleMacro(ByVal bOn As Boolean)
On Error GoTo ErrorHandler

    g_ui.Main.chkEnable.Checked = False
    
    g_Macro.Died = False
    
    If bOn Then
        Call g_Engine.FireStartMacro
    Else
        PrintMessage "Macro Stopped."
        Call g_Macro.StopMacro
    End If
       
Fin:
    Exit Sub
ErrorHandler:
     PrintErrorMessage "ToggleMacro - " & Err.Description
     Resume Fin
End Sub

Public Sub TogglePause()
On Error GoTo ErrorHandler

    g_Macro.Died = False

    If g_Macro.Ticking Then
        If g_Macro.Paused Then
            Call g_Macro.ResumeMacro
            g_ui.Main.btnPause.Text = "Pause"
        Else
            Call g_Macro.PauseMacro
            g_ui.Main.btnPause.Text = "Resume"
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
     PrintErrorMessage "TogglePause - " & Err.Description
     Resume Fin
End Sub

Public Function MonsterEnabled(objMonster As acObject) As Boolean
On Error GoTo ErrorHandler
Dim i As Integer
    
    MonsterEnabled = False

    If g_ui.Macro.chkAttackAny.Checked Then
        MonsterEnabled = True
    ElseIf Valid(objMonster) Then
        If objMonster.UserData(B_ENABLED) Then MonsterEnabled = True
    End If

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "MonsterEnabled - " & Err.Description
    Resume Fin
End Function

Public Function IsPlayer(obj As acObject) As Boolean
    IsPlayer = False
    If Valid(obj) Then IsPlayer = (obj.ObjectType = TYPE_PLAYER)
End Function

Public Function IsMonster(obj As acObject) As Boolean
    IsMonster = False
    If Valid(obj) Then IsMonster = (obj.ObjectType = TYPE_MONSTER)
End Function

Public Function AttackVulnedMobsFirst() As Boolean
    AttackVulnedMobsFirst = g_ui.Macro.chkAttackVulnedOnly.Checked
End Function

Public Function ValidRangeTo(objEntity As acObject, ByVal MaxRange As Single, Optional ByRef fSquareRangeOut As Single) As Boolean
On Error GoTo ErrorHandler

    Dim fRange As Single
    Dim wRange As Single
    
    If Not Valid(objEntity) Then
        PrintErrorMessage "ValidRangeTo - invalid objEntity"
        ValidRangeTo = False
        Exit Function
    End If
    
    'Compare square ranges for speed optimization
    'fSquareRange = objEntity.GetSquareRange
    
    wRange = WorldRange(objEntity.Guid)
    'WorldRange = (g_Filters.g_worldFilter.Distance2D(aGuid, g_Filters.PlayerGUID) * 200)
    
    'Should NEVER have a range of ZERO!!!!
    ' If it does, the means it's out of range
    ' However, it sometimes bugs out, so do some double checking
    If wRange <= 0 Then
        fRange = objEntity.GetRange
        locDebug "ValidRangeTo: WorldRange ZERO:" & fRange & "  for: " & objEntity.Name & " (" & objEntity.Guid & ") LB: " & objEntity.Loc.landblock
    Else
        fRange = wRange
    End If
    
    fSquareRangeOut = fRange
    
    'myDebug "ValidRangeTo: Max:" & MaxRange & " fRange: " & fRange & "  for: " & objEntity.Name
    
    If (fRange <= 0) Then
        ValidRangeTo = False
    Else
        'ValidRangeTo = (CSng(fSquareRange) <= (CSng(MaxRange) * CSng(MaxRange)))
        ValidRangeTo = (fRange <= MaxRange)
    End If
    
    If objEntity.Loc.landblock = 0 Then
        locDebug "ValidRangeTo: Landblock is ZERO: " & objEntity.Name & " (" & objEntity.Guid & ")"
    End If
    
    If wRange <= 0 Or fRange >= 100 Then
        If (objEntity.ObjectType = TYPE_MONSTER) And Not (objEntity.Loc.landblock = g_ds.Player.Loc.landblock) Then
            locDebug "Utils.ValidRangeTo: Monster in Different Landblock: " & objEntity.Name & " : " & objEntity.Guid & " Mlb: " & objEntity.Loc.landblock & " vs Ulb: " & g_ds.Player.Loc.landblock
            'If it's a monster and it has a different landblock than this player, get rid of it in 30 seconds
            'objEntity.canDelete = True
            objEntity.timeData = g_ds.Time + 120
        End If
    End If
    
    'If Not (ValidRangeTo) Then
    '    If (objEntity.ObjectType = TYPE_MONSTER) And Not (landblockInRange(objEntity.Loc.landblock)) Then
    '        MyDebug "Utils.ValidRangeTo: out of range: " & objEntity.Name & " : " & objEntity.Guid
    '        'If it's a monster and it has a different landblock than this player, get rid of it in 30 seconds
    '        objEntity.timeData = g_ds.Time + 30
    '    End If
    'End If
    
    'If Not (ValidRangeTo) Then MyDebug "ValidRangeTo: False! : Max:" & MaxRange & " fRange: " & fRange & "  for: " & objEntity.Name
    
Fin:
    Exit Function
ErrorHandler:
    ValidRangeTo = False
    PrintErrorMessage "ValidRangeTo - " & Err.Description & " (line: " & Erl & ")"
    Resume Fin
End Function


Public Function GetRouteFilePath(ByVal sRouteName As String) As String
    GetRouteFilePath = g_Settings.GetDataFolder & "\" & PATH_ROUTES & "\" & sRouteName & "." & FILE_EXT_ROUTE
End Function

Public Function RouteExist(ByVal sRouteName As String) As Boolean
    RouteExist = FileExists(GetRouteFilePath(sRouteName))
End Function

Public Sub SendTell(ByVal sPlayerName As String, ByVal sMessage As String)
    'Call g_Hooks.SendTellEx(sPlayerName, sMessage)
    g_Core.SendTextToConsole "/t " & sPlayerName & ", " & sMessage, True
End Sub

' SendReplyToConsole is used to talk to +Envoys,
' as /r is only way guranteed to work
Public Sub SendReplyToConsole(ByVal sMessage As String)
    g_Core.SendTextToConsole "/r " & sMessage, True
End Sub

'JSC -- SendMessageByMask
'0x00000800: Fellowship broadcast (@f)
'0x00001000: Patron to vassal (@v)
'0x00002000: Vassal to patron (@p)
'0x00004000: Follower to monarch (@m)
'0x01000000: Covassal broadcast (@c)
'0x02000000: Allegiance broadcast by monarch or speaker (@a)

Public Sub SendFellowshipMessage(ByVal sMessage As String)
    g_Core.SendTextToConsole "/f " & sMessage, True
    'Call g_Hooks.SendMessageByMask(&H800&, sMessage)
End Sub

Public Sub SendAllegianceMessage(ByVal sMessage As String)
    g_Core.SendTextToConsole "/a " & sMessage, True
    'Call g_Hooks.SendMessageByMask(&H2000000, sMessage)
End Sub


Public Function IsEnvoyName(ByVal sPlayerName As String) As Boolean
    sPlayerName = LCase(sPlayerName)
    IsEnvoyName = (InStr(sPlayerName, "+envoy") _
                    Or InStr(sPlayerName, "+") _
                    Or InStr(sPlayerName, "+turbine") _
                    Or InStr(sPlayerName, "envoy"))
End Function


Public Sub RecruitPlayerByGUID(ByVal lPlayerGUID As Long)
On Error GoTo ErrorHandler

    Dim objPlayer As acObject
    Set objPlayer = g_Objects.FindPlayer(lPlayerGUID)
    
    If Valid(objPlayer) Then
        Call g_Macro.RecruitPlayer(objPlayer)
    Else
        MyDebug "RecruitPlayerByGUID - invalid player object"
    End If
       
Fin:
    Set objPlayer = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "RecruitPlayerByGUID - " & Err.Description
    Resume Fin
End Sub

Public Sub GetCurrentSelection(ByRef txtControl As DecalControls.Edit)
On Error GoTo ErrorHandler
    If g_Hooks.CurrentSelection <> 0 Then
        If g_Objects.Items.Exists(g_Hooks.CurrentSelection) Then
            txtControl.Text = g_Objects.Items(g_Hooks.CurrentSelection).Name
        End If
    End If
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "GetCurrentSelection - " & Err.Description
    Resume Fin
End Sub

Public Sub CommonListChange(listControl As DecalControls.list, nX As Long, nY As Long)
    If nX = 1 Then 'delete column
        listControl.DeleteRow (nY)
    End If
End Sub

Public Function ZDiff(ByVal obj1 As acObject, ByVal obj2 As acObject) As Double
On Error GoTo ErrorHandler

    ZDiff = Abs(obj1.Loc.Zoff - obj2.Loc.Zoff)
    
Fin:
    Exit Function
ErrorHandler:
    ZDiff = 0
    PrintErrorMessage "ZDiff - " & Err.Description
    Resume Fin
End Function

Public Function TurboMode() As Boolean
    If Valid(g_ui) Then
        TurboMode = g_ui.Macro.chkEnableTurbo.Checked
    Else
        TurboMode = False
    End If
End Function

Public Function myDateFormat(ByVal aDate As Date) As String
    myDateFormat = Format(aDate, "dddd dd mmm yyyy")
End Function

'Convert the XML file content (the user-interface) to a string
'So it can be stored in ViewShema
Public Function FileToString(sFile As String) As String
On Error GoTo ErrorHandler

  Dim lngFileNr As Long, sLine As String
  
  FileToString = ""
  lngFileNr = FreeFile(0)
  ' Here I'm opening from same directory that scribe.dll is installed,
  ' but you can open from anywhere you'd like.
  Open g_Settings.GetDataFolder & "\" & sFile For Input As #lngFileNr
  Do Until EOF(lngFileNr)
    Line Input #lngFileNr, sLine
    FileToString = FileToString & sLine
  Loop
  Close #lngFileNr
  
Fin:
    Exit Function
ErrorHandler:
    FileToString = ""
    PrintErrorMessage "FileToString(" & sFile & ")"
    Resume Fin
End Function

Public Function IsAdmin(ByVal sSourceName As String) As Boolean
On Error GoTo ErrorHandler

    'Default to false
    IsAdmin = False
    
    If g_ui.Irc.ConnectedToChannel And g_ui.Irc.IsOperator(sSourceName) Then
        IsAdmin = True
    End If
    
    If SameText(sSourceName, g_RemoteCmd.RemoteUserName) And g_RemoteCmd.RemoteAccessON Then
        IsAdmin = True
    End If
    
    'If SameText(sSourceName, "Xeon Xarid") Then
    '    IsAdmin = True
    'End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "IsAdmin(" & sSourceName & ") - " & Err.Description
    IsAdmin = False
    Resume Fin
End Function


Public Function MagicSchoolToSkillId(ByVal iMagicSchool As Integer) As Long
Dim lRet As Long

    Select Case iMagicSchool
        Case SCHOOL_CREATURE
            lRet = eSkillCreatureEnchantment
        Case SCHOOL_ITEM
            lRet = eSkillItemEnchantment
        Case SCHOOL_WAR
            lRet = eSkillWarMagic
        Case SCHOOL_LIFE
            lRet = eSkillLifeMagic
        Case Else
            PrintErrorMessage "MagicSchoolToSkillId: unknown magic school " & iMagicSchool
            lRet = eSkillCreatureEnchantment
    End Select
    
    MagicSchoolToSkillId = lRet
End Function

Public Function MagicSchoolIsTrained(ByVal iSchoolId As eMagicSchools) As Boolean
    MagicSchoolIsTrained = (g_Hooks.Skill(MagicSchoolToSkillId(iSchoolId)) > 0)
End Function

Public Function HasWarMagic() As Boolean
    HasWarMagic = MagicSchoolIsTrained(SCHOOL_WAR)
End Function

Public Sub SortCollection(ByRef ColVar As Collection)
On Error GoTo ErrorHandler
    
    Dim oCol As Collection
    Dim i As Integer
    Dim i2 As Integer
    Dim iBefore As Integer
    
    If ColVar.Count > 0 Then
        Set oCol = New Collection
        For i = 1 To ColVar.Count
            If oCol.Count = 0 Then
                oCol.Add ColVar(i)
            Else
                iBefore = 0
                For i2 = oCol.Count To 1 Step -1
                    If LCase(ColVar(i)) < LCase(oCol(i2)) Then
                        iBefore = i2
                    Else
                        Exit For
                    End If
                Next
                    
                If iBefore = 0 Then
                    oCol.Add ColVar(i)
                Else
                    oCol.Add ColVar(i), , iBefore
                End If
            End If
        Next
        
        Set ColVar = oCol
        Set oCol = Nothing
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SortCollection - " & Err.Description
    Resume Fin
End Sub

Sub SortByWorkmanship(ItemsList() As acObject, Optional ByVal bDescendant = True)
On Error GoTo ErrorHandler

    Dim i As Integer
    Dim j As Integer
    Dim best_item As acObject
    Dim best_j As Integer
    Dim bDoSwap As Boolean
    Dim max As Integer, min As Integer
    
    max = UBound(ItemsList)
    min = LBound(ItemsList)

    For i = min To max - 1
        Set best_item = ItemsList(i)
        best_j = i
        For j = i + 1 To max
            If bDescendant Then
                bDoSwap = (ItemsList(j).Workmanship > best_item.Workmanship)
            Else
                bDoSwap = (ItemsList(j).Workmanship < best_item.Workmanship)
            End If
            
            If bDoSwap Then
                Set best_item = ItemsList(j)
                best_j = j
            End If
        Next j
        Set ItemsList(best_j) = ItemsList(i)
        Set ItemsList(i) = best_item
    Next i
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SortByWorkmanship - " & Err.Description
    Resume Fin
End Sub

'For salvage bags, sort by quantity of salvage units currently in bag
Public Sub SortBySalvageQuantity(ItemsList() As acObject, Optional ByVal bDescendant = True)
On Error GoTo ErrorHandler

    Dim i As Integer
    Dim j As Integer
    Dim best_item As acObject
    Dim best_j As Integer
    Dim bDoSwap As Boolean
    Dim max As Integer, min As Integer
    
    max = UBound(ItemsList)
    min = LBound(ItemsList)

    For i = min To max - 1
        Set best_item = ItemsList(i)
        best_j = i
        For j = i + 1 To max
            If bDescendant Then
                bDoSwap = (ItemsList(j).UsesLeft > best_item.UsesLeft)
            Else
                bDoSwap = (ItemsList(j).UsesLeft < best_item.UsesLeft)
            End If
            
            If bDoSwap Then
                Set best_item = ItemsList(j)
                best_j = j
            End If
        Next j
        Set ItemsList(best_j) = ItemsList(i)
        Set ItemsList(i) = best_item
    Next i
    
Fin:
    Set best_item = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SortBySalvageQuantity - " & Err.Description
    Resume Fin
End Sub

'Returns an array filled with all objects from collection (0 to colObj.count - 1)
'Collection isnt altered
Public Function ColToArray(colObj As colObjects) As acObject()
On Error GoTo ErrorHandler

    Dim obj As acObject
    Dim i As Integer
    Dim theArray() As acObject
    
    i = 0
    ReDim theArray(0 To colObj.Count - 1)
    For Each obj In colObj
        Set theArray(i) = obj
        i = i + 1
    Next obj
    
    ColToArray = theArray
    
Fin:
    Set obj = Nothing
    Exit Function
ErrorHandler:
    PrintErrorMessage "ColToArray - " & Err.Description
    Resume Fin
End Function

Public Sub ClickSalvageButton()
On Error GoTo ErrorHandler

    'Dim x As Integer, y As Integer
    'x = g_Hooks.Area3DWidth - 65
    'y = g_Hooks.Area3DHeight + 85 'takes health bar height at top into account
    'MyDebug "Clicking Salvage Button at @ (" & x & "," & y & ")"
    'Call g_Core.MouseClick(x, y)
    
    Call g_Hooks.SalvagePanelSalvage
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ClickSalvageButton - " & Err.Description
    Resume Fin
End Sub

Public Sub ClickVendorButton()
On Error GoTo ErrorHandler

    'MyDebug "ClickVendorButton: Top: " & g_Hooks.AC3DRegionRect.Top & " - Bottom: " & g_Hooks.AC3DRegionRect.Bottom
    'MyDebug "ClickVendorButton: Left: " & g_Hooks.AC3DRegionRect.Left & " - Right: " & g_Hooks.AC3DRegionRect.Right

    Dim x As Integer, y As Integer
    x = g_Hooks.AC3DRegionRect.Right - 40
    y = g_Hooks.AC3DRegionRect.Bottom + 60
    MyDebug "Clicking Vendor Button at @ (" & x & "," & y & ")"
    Call g_Core.MouseClick(x, y)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ClickVendorButton - " & Err.Description
    Resume Fin
End Sub

Public Sub ClickVendorItemsTab()
On Error GoTo ErrorHandler

    Dim x As Integer, y As Integer
    x = 35
    y = g_Hooks.AC3DRegionRect.Bottom + 30
    MyDebug "Clicking Items Tab at @ (" & x & "," & y & ")"
    Call g_Core.MouseClick(x, y)
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ClickVendorItemsTab - " & Err.Description
    Resume Fin
End Sub

Public Sub ClickButton(xOff As Long, yOff As Long, Optional ByVal useOffSet As Boolean = True)
    Dim iPosX As Integer, iPosY As Integer
    
    Call g_Window.UpdateDimensions
    'iPosX = g_Hooks.Area3DWidth + Xoff
    'iPosY = g_Window.Height - Yoff
    
    'MyDebug "ClickButton: Top: " & g_Hooks.AC3DRegionRect.Top & " - Bottom: " & g_Hooks.AC3DRegionRect.Bottom
    'MyDebug "ClickButton: Left: " & g_Hooks.AC3DRegionRect.Left & " - Right: " & g_Hooks.AC3DRegionRect.Right

    MyDebug "Utils.ClickButton: g_Window.Width: " & g_Window.Width & " - g_Window.Height " & g_Window.Height

    iPosX = g_Window.Width - xOff
    iPosY = g_Window.Height - yOff
    
    MyDebug "Utils.ClickButton: Xoff: " & xOff & "   Yoff: " & yOff
    MyDebug "Utils.ClickButton: iPosX: " & iPosX & "  iPosY: " & iPosY
    
    If useOffSet Then
        Call g_Core.MouseClick(iPosX, iPosY)
    Else
        Call g_Core.MouseClick(xOff, yOff)
    End If
    
End Sub


'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[Utils] " & DebugMsg, bSilent)
    End If
End Sub


