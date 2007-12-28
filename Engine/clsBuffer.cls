VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEBUG_ME = False

Private m_BuffQueue As clsSpellQueue        'list of pending buff spells
Private m_colBuffs As Collection            'list of all the spells the buffer has to use
Private m_existBuffs As Collection          'list of existing buffs and their times
Private m_iRebuffMode As Integer

Private WithEvents m_tmrNextRebuff As clsTimer
Attribute m_tmrNextRebuff.VB_VarHelpID = -1
Private m_dRebuffInterval As Double         'time in seconds
Private m_dBuffCycleLen As Double           'time for a full buff cycle, in seconds
Private m_bRepeat As Boolean                'repeat buff cycles?

Private m_sCurSpellName As String           'Current spell name

'Continuous Rebuff Mode
Private m_lCurBuff As Long      'index of the current buff
Private m_iNumBuffs As Integer  'number of consecutive buffs to cast in continuous buff mode

'Events
Public Event OnCycleComplete()

'Public
Public mNumBuffsToCast As Integer


'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################
Private Sub Class_Initialize()
    Set m_tmrNextRebuff = CreateTimer
    
    'default rebuff mode
    m_iRebuffMode = REBUFF_FULL
    m_dRebuffInterval = 10          'time in seconds
    m_iNumBuffs = 5
    m_bRepeat = True
    mNumBuffsToCast = 0
    
    Call Reset
End Sub

Private Sub Reset()
    Set m_BuffQueue = New clsSpellQueue
    Set m_colBuffs = New Collection
    Call m_tmrNextRebuff.Reset
    m_lCurBuff = -1
    m_sCurSpellName = ""
End Sub


Private Sub Class_Terminate()
    Set m_BuffQueue = Nothing
    Set m_colBuffs = Nothing
    Set m_tmrNextRebuff = Nothing
End Sub

'#####################################################################################
'#
'#                               PROPERTIES
'#
'#####################################################################################

Public Property Get BuffQueue() As clsSpellQueue
    Set BuffQueue = m_BuffQueue
End Property

Public Property Get RebuffMode() As Integer
    RebuffMode = m_iRebuffMode
End Property

Public Property Let RebuffMode(ByVal iMode As Integer)
    m_iRebuffMode = iMode
End Property

Public Property Get BuffCycleLen() As Double        'time in seconds
    BuffCycleLen = m_dBuffCycleLen
End Property

Public Property Let BuffCycleLen(ByVal dVal As Double)  'time in seconds
    m_dBuffCycleLen = dVal
    'update rebuff interval
    m_dRebuffInterval = CalcRebuffInterval
End Property

Public Property Get NumContBuffs() As Integer
    NumContBuffs = m_iNumBuffs
End Property

Public Property Let NumContBuffs(ByVal iVal As Integer)
    m_iNumBuffs = iVal
    'update rebuff interval
    m_dRebuffInterval = CalcRebuffInterval
End Property

Public Property Get RebuffInterval() As Integer
    RebuffInterval = m_dRebuffInterval
End Property

Public Property Get RepeatCycles() As Boolean
    RepeatCycles = m_bRepeat
End Property

Public Property Let RepeatCycles(ByVal bVal As Boolean)
    m_bRepeat = bVal
End Property

Public Property Get NextRebuff() As clsTimer
    Set NextRebuff = m_tmrNextRebuff
End Property

'#####################################################################################
'#
'#                                PRIVATE
'#
'#####################################################################################

Private Sub AddBuffToList(Optional ByVal sFamily As String = "", _
                    Optional ByVal iType As Integer = SPELLTYPE_NORMAL, _
                    Optional ByVal iElement As Integer = DMG_NONE, _
                    Optional objTarget As acObject = Nothing, _
                    Optional ByVal iLevelWanted = 7, _
                    Optional ByVal iSchool As Integer = SCHOOL_CREATURE)
On Error GoTo ErrorHandler

    Dim objBuff As New clsSpellQueueItem
    With objBuff
        .Index = m_colBuffs.Count
        .SpellFamily = sFamily
        .SpellType = iType
        .SpellElement = iElement
        .SpellSchool = iSchool
        
        If Valid(objTarget) Then
            .TargetGUID = objTarget.Guid
            .TargetName = objTarget.Name
        Else
            .TargetName = "Self"
            .TargetGUID = 0
        End If
        
    End With
    
    Call m_colBuffs.Add(objBuff, CStr(objBuff.Index))
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "AddBuffToList"
    Resume Fin
End Sub

Private Function FindBuff(ByVal Index As Long, ByRef objBuffOut As clsSpellQueueItem) As Boolean
On Error GoTo NotFound
    Dim bRet As Boolean

    Set objBuffOut = m_colBuffs(CStr(Index))
    bRet = True
Fin:
    FindBuff = bRet
    Exit Function
NotFound:
    bRet = False
    Resume Fin
End Function


Private Sub CheckItemBuff(ByVal bChecked As Boolean, ByVal sSpellFamily As String, objItem As acObject)
On Error GoTo ErrorHandler
    Dim objBuff As clsSpellQueueItem
    
    If bChecked And Valid(objItem) Then
        Call AddBuffToList(sSpellFamily, , , objItem, , SCHOOL_ITEM)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckItemBuff"
    Resume Fin
End Sub

Private Sub CheckItemBanes(ByVal bChecked As Boolean, objItem As acObject)
    If bChecked And Valid(objItem) Then
        Call AddBanes(objItem)
    End If
End Sub

Private Sub CheckCollectionBanes(ByVal bChecked As Boolean, colItems As colObjects)
On Error GoTo ErrorHandler
    
    Dim objItem As acObject

    If bChecked Then
        For Each objItem In colItems
            Call AddBanes(objItem)
        Next objItem
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckCollectionBanes"
    Resume Fin
End Sub

Private Sub CheckAddBane(ByVal iElement As Integer, ByVal bChecked As Boolean, objItem As acObject)
    If bChecked Then
        Call AddBuffToList(, SPELLTYPE_BANE, iElement, objItem, , SCHOOL_ITEM)
    End If
End Sub


Private Sub AddBanes(objItem As acObject)
On Error GoTo ErrorHandler

    If Not Valid(objItem) Then
        PrintErrorMessage "clsBuffer.AddBanes : invalid objItem"
        Exit Sub
    End If
    
    With g_ui.Buffs
        Call CheckAddBane(DMG_NONE, .chkItmBaneImpen.Checked, objItem)
        Call CheckAddBane(DMG_SLASHING, .chkItmBaneSlash.Checked, objItem)
        Call CheckAddBane(DMG_PIERCING, .chkItmBanePierce.Checked, objItem)
        Call CheckAddBane(DMG_BLUDGEONING, .chkItmBaneBludg.Checked, objItem)
        Call CheckAddBane(DMG_FIRE, .chkItmBaneFire.Checked, objItem)
        Call CheckAddBane(DMG_COLD, .chkItmBaneFrost.Checked, objItem)
        Call CheckAddBane(DMG_ACID, .chkItmBaneAcid.Checked, objItem)
        Call CheckAddBane(DMG_LIGHTNING, .chkItmBaneLightning.Checked, objItem)
    End With

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "AddBanes"
    Resume Fin
End Sub

Public Sub ForceRestart()
    Call m_tmrNextRebuff_OnTimeout
    m_tmrNextRebuff.Enabled = False
End Sub

'Returns the effective rebuff interval, in seconds, depending on the rebuff mode
Private Function CalcRebuffInterval() As Double
On Error GoTo ErrorHandler
    Dim dRet As Double
    
    Select Case m_iRebuffMode
        Case REBUFF_CONTINUOUS
            Dim cnt As Long
            Dim iNumPacks As Integer
            Dim iBuffsInLastPack As Integer
            
            cnt = m_colBuffs.Count
            
            If cnt <= 0 Then cnt = 1
            If m_iNumBuffs < 1 Then m_iNumBuffs = 1
        
            iBuffsInLastPack = cnt Mod m_iNumBuffs
            iNumPacks = (cnt + (m_iNumBuffs - iBuffsInLastPack)) / m_iNumBuffs
            If iNumPacks < 1 Then iNumPacks = 1
    
            'compute the continuous rebuff interval
            dRet = m_dBuffCycleLen / CDbl(iNumPacks)
            MyDebug "CalcRebuffInterval - BuffsCnt=" & cnt & ", iBuffsInLastPack=" & iBuffsInLastPack & ", iNumPacks=" & iNumPacks & " -- Interval:" & myFormatTime(dRet)
            
        Case Else
            dRet = m_dBuffCycleLen
    End Select
    
Fin:
    CalcRebuffInterval = dRet
    Exit Function
ErrorHandler:
    dRet = -1
    PrintErrorMessage "CalcRebuffInterval"
    Resume Fin
End Function

'Called when starting a new buff cycle
Private Sub OnRestartBuffCycle()
    'rebuild the bufflist
    Call BuildBuffList
    
    m_lCurBuff = -1
End Sub

Private Sub FillBuffQueue(ByVal iFillMode As Integer, Optional ByVal bCleanQueue As Boolean = False)
On Error GoTo ErrorHandler

    Dim objBuff As clsSpellQueueItem
    
    If iFillMode = REBUFF_CONTINUOUS Then
        Dim i As Integer
        
        For i = 1 To m_iNumBuffs
            m_lCurBuff = m_lCurBuff + 1
            
            If m_lCurBuff > m_colBuffs.Count - 1 Then  ' count - 1 ?
                MyDebug "clsBuffer - Cont Buff Cycle over"
                Call OnRestartBuffCycle
                GoTo Fin
            End If
            
            If FindBuff(m_lCurBuff, objBuff) Then
                Call m_BuffQueue.Add(objBuff)
            Else
                PrintWarning "FillBuffQueue - CONTINUOUS - Couldn't find Buff #" & m_lCurBuff
            End If
        Next i
    
    Else 'REBUFF_FULL
        
        'clean up for new buff cycle
        Call OnRestartBuffCycle
        
        'copy all the spells from the buff list to the pending buffs queue
        For Each objBuff In m_colBuffs
            Call m_BuffQueue.Add(objBuff)
        Next objBuff
        
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "AddBanes"
    Resume Fin
End Sub

Private Sub PrepareNextRebuffTime()
    'setup timer
    m_tmrNextRebuff.SetNextTime m_dRebuffInterval
End Sub

Private Sub m_tmrNextRebuff_OnTimeout()
    MyDebug "tmrNextRebuff Timeout"

    If m_bRepeat Then
        'push spells to the buff queue
        Call PushBuffs
    
        'next rebuff
        Call PrepareNextRebuffTime
    End If
End Sub

'#####################################################################################
'#
'#                                PUBLIC
'#
'#####################################################################################

Public Sub BuildBuffList()
On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim sSpellFam As String
    Dim bChecked As Boolean
    Dim objItem As acObject
    
    Call g_Objects.Equipment.Update
    
    Set m_colBuffs = New Collection
    With g_ui.Buffs
    
        'Buff Wand 1st
        CheckItemBuff .chkItmWandHL.Checked, SPELL_HERMETIC_LINK, g_Data.Wand
        CheckItemBuff .chkItmWandDef.Checked, SPELL_DEF, g_Data.Wand
        CheckItemBuff .chkItmWandSD.Checked, SPELL_SPIRIT_DRINKER, g_Data.Wand
        
        If .chkEnableCreatureBuffs.Checked Then
            For i = 0 To .lstCreature.Count - 1
                bChecked = .lstCreature.Data(2, i, 0)
                sSpellFam = .lstCreature.Data(1, i, 0)
                If bChecked Then
                    Call AddBuffToList(sSpellFam)
                End If
            Next i
        End If
        
        If .chkEnableLifeBuffs.Checked Then
            For i = 0 To .lstLifePros.Count - 1
                bChecked = .lstLifePros.Data(2, i, 0)
                sSpellFam = .lstLifePros.Data(1, i, 0)
                If bChecked Then
                    Call AddBuffToList(sSpellFam)
                End If
            Next i
        End If
        
        'Armor Banes
        'SPK -  Added option to use the new item banes system, where you only have to select yourself
        '       to cast the spells
        If .chkItmSelectSelf.Checked Then
            CheckItemBanes .chkItmSelectSelf.Checked, g_Objects.Player
        Else
            CheckItemBanes .chkItmArmorHead.Checked, g_Objects.Equipment.head
            CheckItemBanes .chkItmArmorHands.Checked, g_Objects.Equipment.Hands
            CheckItemBanes .chkItmArmorFeet.Checked, g_Objects.Equipment.Feet
            CheckItemBanes .chkItmArmorTopUndie.Checked, g_Objects.Equipment.TopUndie
            CheckItemBanes .chkItmArmorBottomUndie.Checked, g_Objects.Equipment.BottomUndie
            CheckCollectionBanes .chkItmArmorTop.Checked, g_Objects.Equipment.Top
            CheckCollectionBanes .chkItmArmorBottom.Checked, g_Objects.Equipment.Bottom
        End If
        
        '*****************************************************************************
        'Special macro items which may not be currently equipped but need to be buffed
        '*****************************************************************************
        
        'Weapon / Bow / Shield
        If IsMelee Then
            Dim objWeapon As acObject
            
            If g_Macro.CombatType = TYPE_MELEE Then
                Set objWeapon = g_Data.Weapon
                CheckItemBanes .chkItmArmorShield.Checked, g_Data.Shield
                CheckItemBuff .chkItmWeapHS.Checked, SPELL_HS, objWeapon
            Else    'Archer
                Set objWeapon = g_Data.Bow
            End If
            
            CheckItemBuff .chkItmWeapBD.Checked, SPELL_BD, objWeapon
            CheckItemBuff .chkItmWeapSK.Checked, SPELL_SK, objWeapon
            CheckItemBuff .chkItmWeapDE.Checked, SPELL_DEF, objWeapon
        End If

    End With

    MyDebug "BuildBuffList: total spells: " & m_colBuffs.Count
    
    Call existingSpellList

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "BuildBuffList - " & Err.Description
    Resume Fin
End Sub

Public Sub existingSpellList()
On Error GoTo ErrorHandler
    
    Dim newSpell As New clsSpellQueueItem
    Dim oldSpell As New clsSpellQueueItem
    Dim isFound As Boolean
    Dim i As Long
    Dim iSpell As Long

    MyDebug "Displaying Existing Spells: " & g_Filters.ActiveSpellsCount

    Set m_existBuffs = New Collection
    
    For i = 0 To (g_Filters.ActiveSpellsCount - 1)
        'g_Filters.ActiveSpell(i).SpellID
        'g_Filters.ActiveSpell(i).Family
        'g_Filters.ActiveSpell(i).Layer
        'g_Filters.ActiveSpell(i).TimeRemaining

        Set newSpell = New clsSpellQueueItem
        
        newSpell.SpellID = g_Filters.ActiveSpell(i).SpellID
        newSpell.SpellFamily = g_Filters.ActiveSpell(i).Family
        newSpell.TimeRemaining = g_Filters.ActiveSpell(i).TimeRemaining
        
        If m_existBuffs.Count = 0 Then
            Call m_existBuffs.Add(newSpell)
        End If
        
        isFound = False
        
        For iSpell = 1 To m_existBuffs.Count
             Set oldSpell = m_existBuffs.Item(iSpell)
                If (oldSpell.SpellFamily = newSpell.SpellFamily) Then
                    isFound = True
                    If (oldSpell.TimeRemaining < newSpell.TimeRemaining) Then
                        Call m_existBuffs.Add(newSpell, , , iSpell) 'add newSpell after existing one
                        Call m_existBuffs.Remove(iSpell) ' remove existing one
                    End If
                End If
        Next iSpell
        
        If Not isFound Then
            Call m_existBuffs.Add(newSpell)
        End If
        
        'MyDebug "ID: " & g_Filters.ActiveSpell(i).SpellID & " :F: " & g_Filters.ActiveSpell(i).Family & " :T: " & g_Filters.ActiveSpell(i).TimeRemaining & " :D: " & g_Filters.ActiveSpell(i).Duration
        'MyDebug "NID: " & newSpell.SpellID & " :F: " & newSpell.SpellFamily & " :T: " & newSpell.TimeRemaining
    
    Next i

    'For iSpell = 1 To m_existBuffs.Count
    '    Set oldSpell = m_existBuffs.Item(iSpell)
    '    MyDebug "SpellID: " & oldSpell.SpellID & " :F: " & oldSpell.SpellFamily & " :T: " & oldSpell.TimeRemaining
    'Next iSpell

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "existingSpellList - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Public Sub DebugList()
    MyDebug "Displaying BuffList"
    Dim sOut As String
    Dim objBuff As clsSpellQueueItem
    For Each objBuff In m_colBuffs
        sOut = objBuff.Description & " - Target: "
        If objBuff.TargetGUID <> 0 Then
            sOut = sOut & objBuff.TargetName
        Else
            sOut = sOut & "Self"
        End If
        MyDebug sOut
    Next objBuff
    
    MyDebug "-------- Total : " & m_colBuffs.Count
    
End Sub

'If bForceFullMode then all the buffs from the buff list will be pushed to the buff queue
'else, it depends of the current buff mode (full or continuous/partial)
'if bCleanQueue then the buff queue content will be cleared before buffs get pushed
Public Sub PushBuffs(Optional ByVal bForceFullRebuff As Boolean = False, Optional ByVal bCleanQueue As Boolean = False)
    If bForceFullRebuff Then
        Call FillBuffQueue(REBUFF_FULL, bCleanQueue)
    Else
        Call FillBuffQueue(m_iRebuffMode, bCleanQueue)
    End If
End Sub

Public Sub StartService(ByVal iBuffMode As Integer, Optional ByVal iRebuffInterval_min As Integer = 45, Optional ByVal bRepeatCycles As Boolean = True, Optional ByVal bPushBuffsNow As Boolean = True, Optional ByVal iContinuousBuffs = 1)
On Error GoTo ErrorHandler

    MyDebug "clsBuffer.StartService(" & iBuffMode & ", " & CStr(bRepeatCycles) & ", " & iContinuousBuffs & ")"
    
    'first reset all
    Call Reset
    
    'build the buffs list
    Call BuildBuffList
    
    'Setup params
    m_iRebuffMode = iBuffMode
    m_bRepeat = bRepeatCycles
    m_iNumBuffs = iContinuousBuffs
    
    'Set the length of a complete buff cycle
    BuffCycleLen = CDbl(iRebuffInterval_min * 60)

    'push buffs to the queue if required
    If bPushBuffsNow Then
        MyDebug "StartService - Pushing buffs to list"
        Call PushBuffs(False, True)
    End If
    
    'setup next rebuff time if required
    If m_bRepeat Then
        Call PrepareNextRebuffTime
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsBuffer.StartService"
    Resume Fin
End Sub

Public Sub StopService()
    MyDebug "clsBuffer.StopService()"
    Call Reset
End Sub

'Casts the next spell in the buff queue
'--> However, it doesnt remove it from top of queue (this is done in OnReady event after casting)
Public Sub CastNextSpell()
On Error GoTo ErrorHandler

    Dim objBuff As clsSpellQueueItem
    Dim objSpell As clsSpell
    Dim objItem As acObject                'pointer to the object to cast on
    Dim spellList As clsSpellList
    
    Dim lTargetGUID As Long
    Dim sTargetName As String
    
    If m_BuffQueue.Count <= 0 Then
        Exit Sub
    End If
    
    Set objBuff = m_BuffQueue.ReadTop
    
    'Item or Creature/Life buff ?
    If objBuff.SpellSchool = SCHOOL_ITEM Then
    
        'choose the right spell database
        Set spellList = g_Spells.Items
        
        'find the object we want to cast on
        Set objItem = g_Objects.FindObject(objBuff.TargetGUID)
        
        'Make sure we found an item to buff (should always be the case...)
        If Not Valid(objItem) Then
            PrintWarning "clsBuffer.CastNextSpell : couldn't find the target object to buff. Ignoring"
            Call m_BuffQueue.Pop    'remove the bad buff from queue
            GoTo Fin
        End If
        
        'Setup the casting params
        lTargetGUID = objItem.Guid
        sTargetName = objItem.Name
    
    Else    'Creature/Life buff
    
        'choose the right spell database
        Set spellList = g_Spells.sBuffs
        
        'Setup the casting params
        lTargetGUID = objBuff.TargetGUID
        sTargetName = objBuff.TargetName
        
    End If
    
    locDebug "clsBuffer.NextSpell:Spell level wanted : " & objBuff.LevelWanted
    
    'Find the spell infos
    If objBuff.SpellType <> SPELLTYPE_NORMAL Then   'if it's a typed spell (i.e banes)
        Set objSpell = spellList.FindSpellByType(objBuff.SpellType, objBuff.SpellElement, g_Data.BuffsSpellsLevel)
        locDebug "clsBuffer.NextSpell:Typed Spell: wanted: " & objBuff.SpellType
    Else
        Set objSpell = spellList.FindSpell(objBuff.SpellFamily, g_Data.BuffsSpellsLevel)
        locDebug "clsBuffer.NextSpell: SpellFamily: " & objBuff.SpellFamily
    End If
    
    'Check if we have been able to find the spell, and cast it
    If Not Valid(objSpell) Then
        PrintErrorMessage "Could not find a valid spell for " & objBuff.Description & " - Igoring it."
        Call m_BuffQueue.Pop    'remove the bad buff from queue
        GoTo Fin
    Else
        If Not (IsBusy) Then ' Prevent spell spamming
        
            locDebug "clsBuffer.CastNextSpell: casting: " & objSpell.SpellName
            
            m_sCurSpellName = objSpell.SpellName
            
            If Not (g_ui.Options.chkFilterLTMsg.Checked) Then
                If lTargetGUID = g_Objects.Player.Guid Or lTargetGUID = 0 Then
                    PrintMessage "[Rebuff] Casting " & objSpell.SpellName
                Else
                    PrintMessage "[Rebuff] Casting " & objSpell.SpellName & " on " & sTargetName
                End If
            End If
        
            Call g_Spells.CastThisSpell(objSpell, lTargetGUID)
        Else
            locDebug "clsBuffer.CastNextSpell: IsBusy TRUE, _not_ casting: " & objSpell.SpellName
        End If
    End If
    
    
Fin:
    Set spellList = Nothing
    Set objItem = Nothing
    Set objBuff = Nothing
    Set objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CastNextSpell"
    Resume Fin
End Sub

'Look for "has expired" messages to make sure we don't drop any buffs
Public Sub CheckExpiredSpell(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    Dim sSpellName As String
    Dim objBuff As clsSpellQueueItem
    Dim objSpell As clsSpell
    Dim i As Long
    
'   The spell Cragstone's Will on Opal Blunt Baton has expired.
'   Adja's Blessing has Expired
    
    sSpellName = Replace(sCastMsg, " has expired.", "")
    MyDebug "CheckExpiredSpell: " & sSpellName
    
    Set objSpell = g_Spells.sBuffs.FindSpellByName(sSpellName)
    
    If Not Valid(objSpell) Then
        ' Can't find this spell anywhere
        MyDebug "CheckExpiredSpell: not a valid spell name: " & sSpellName
        Exit Sub
    End If

    'First, check to see if we have this spell still active on us
    For i = 0 To (g_Filters.ActiveSpellsCount - 1)
        
        If g_Filters.ActiveSpell(i).SpellID = objSpell.SpellID Then
            ' Yep, found active spell, so we are A-OK!
            MyDebug "CheckExpiredSpell: found active spell!"
            Exit Sub
        End If
    
    Next i
    
    'If not then we need to see if it's in the Buff Queue
    
    For i = m_lCurBuff To m_iNumBuffs
            
        If FindBuff(i, objBuff) Then
            If objBuff.SpellID = objSpell.SpellID Then
                Call m_BuffQueue.Add(objBuff)
                MyDebug "CheckExpiredSpell: found in m_colBuffs, adding to current Queue"
                Exit Sub
            End If
        End If
    Next i
    
    'For Each objBuff In m_colBuffs
    '    'objBuff.Description
    '    If objBuff.SpellID = objSpell.SpellID Then
    '    End If
    'Next objBuff

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckExpiredSpell - " & Err.Description
    Resume Fin
End Sub
        
        
'Read the AC Console to figure out if we have cast the current buff spell, if so pop it from the queue
'ex: You cast Adja's Blessing you yourself
Public Sub CheckCastedSpell(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    If m_sCurSpellName <> "" Then
        If InStr(LCase(sCastMsg), LCase(m_sCurSpellName)) Then
            
            locDebug "clsBuffer.CheckCastedSpell - Current spell " & m_sCurSpellName & " has been cast"
            
            If (g_Macro.State = ST_REBUFF) Then
                MyDebug "clsBuffer.CheckCastedSpell: poping m_sCurSpellName off BuffQueue"
                Call m_BuffQueue.Pop
                Call g_Spells.c_SpellQueue.Pop
                'update progress bar
                g_ui.Main.progBuffs.Value = g_ui.Main.progBuffs.Value + 1
                m_sCurSpellName = ""
            End If
            
            If g_Spells.Casting Then
                locDebug "clsBuffer.CheckCastedSpell - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        Else
            locDebug "clsBuffer.CheckCastedSpell doesn't match: " & m_sCurSpellName & " : " & sCastMsg
            If g_Spells.Casting Then
                locDebug "clsBuffer.CheckCastedSpell - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        End If
    Else
        locDebug "clsBuffer.CheckCastedSpell: m_sCurSpellName is Blank"
        If g_Spells.Casting Or Not g_Macro.OkToCast Then
            locDebug "clsBuffer.CheckCastedSpell - g_Spells.Casting is true, calling OnSpellCastComplete"
            Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
        End If

    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckCastedSpell - " & Err.Description
    Resume Fin
End Sub

'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsBuffer] " & DebugMsg, bSilent)
    End If
End Sub


