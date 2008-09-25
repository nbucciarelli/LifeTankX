VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuddyBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEBUG_ME = False

Private Const SECURE_TIMER = 10

Private m_BuffQueue As clsSpellQueue        'list of pending buff spells
Private m_colBuffs As Collection            'list of all the spells the buffer has to use
Private m_iRebuffMode As Integer

Private WithEvents m_tmrNextRebuff As clsTimer
Attribute m_tmrNextRebuff.VB_VarHelpID = -1
Private m_dRebuffInterval As Double         'time in seconds
Private m_dBuffCycleLen As Double           'time for a full buff cycle, in seconds
Private m_bRepeat As Boolean                'repeat buff cycles?

Private m_sCurSpellName As String           'Current spell name
Private m_sCurTargetName As String          'Current Buff Item target name

'Continuous Rebuff Mode
Private m_lCurBuff As Long      'index of the current buff
Private m_iNumBuffs As Integer  'number of consecutive buffs to cast in continuous buff mode

'Buff Buddy held items
Private m_HeldItem As acObject
Private m_HeldShield As acObject

Private m_tmrSecure As clsTimer

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
    Set m_HeldItem = Nothing
    Set m_HeldShield = Nothing
    Set m_tmrSecure = CreateTimer

    Call Reset
End Sub

Private Sub Reset()
    Set m_BuffQueue = New clsSpellQueue
    Set m_colBuffs = New Collection
    Call m_tmrNextRebuff.Reset
    Call m_tmrSecure.Reset
    m_lCurBuff = -1
    m_sCurSpellName = ""
    m_sCurTargetName = ""
End Sub


Private Sub Class_Terminate()
    Set m_BuffQueue = Nothing
    Set m_colBuffs = Nothing
    Set m_tmrNextRebuff = Nothing
    Set m_tmrSecure = Nothing
    Set m_HeldItem = Nothing
    Set m_HeldShield = Nothing
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

Public Sub setHeldItem(ByVal anObj As acObject)
    Set m_HeldItem = anObj
End Sub

Public Sub setHeldShield(ByVal anObj As acObject)
    Set m_HeldShield = anObj
End Sub

'#####################################################################################
'#
'#                                PRIVATE
'#
'#####################################################################################

Private Sub AddBuffToList(ByVal objTarget As acObject, _
                    Optional ByVal sFamily As String = "", _
                    Optional ByVal iType As Integer = SPELLTYPE_NORMAL, _
                    Optional ByVal iElement As Integer = DMG_NONE, _
                    Optional ByVal iLevelWanted = 8, _
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
            MyDebug "clsBuddyBuffer.addBuffToList: ObjTarget: " & objTarget.Name
            .TargetGUID = objTarget.Guid
            .TargetName = objTarget.Name
        Else
            MyDebug "clsBuddyBuffer.addBuffToList: !Valid ObjTarget "
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
    
    If Not Valid(objItem) Then Exit Sub
    If objItem.unEnchantable Then Exit Sub
    
    If bChecked And Valid(objItem) Then
        Call AddBuffToList(objItem, sSpellFamily, , , , SCHOOL_ITEM)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckItemBuff"
    Resume Fin
End Sub

Private Sub CheckItemBanes(ByVal bChecked As Boolean, objItem As acObject)
    If Not Valid(objItem) Then Exit Sub
    If objItem.unEnchantable Then Exit Sub
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
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckCollectionBanes"
    Resume Fin
End Sub

Private Sub CheckAddBane(ByVal iElement As Integer, ByVal bChecked As Boolean, objItem As acObject)
    If Not Valid(objItem) Then Exit Sub
    If objItem.unEnchantable Then Exit Sub
    If bChecked Then
        Call AddBuffToList(objItem, SPELLTYPE_BANE, iElement, , SCHOOL_ITEM)
    End If
End Sub


Private Sub AddBanes(objItem As acObject)
On Error GoTo ErrorHandler

    If Not Valid(objItem) Then Exit Sub
    If objItem.unEnchantable Then Exit Sub
    
    If Not Valid(objItem) Then
        PrintErrorMessage "clsBuddyBuffer.AddBanes : invalid objItem"
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
    Call m_tmrSecure.Reset
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
            
            If m_iNumBuffs > 0 Then
                iNumPacks = (cnt + (m_iNumBuffs - iBuffsInLastPack)) / m_iNumBuffs
            End If
            
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
                MyDebug "clsBuddyBuffer - Cont Buff Cycle over"
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
    Call m_tmrNextRebuff.SetNextTime(m_dRebuffInterval)
End Sub

Private Sub m_tmrNextRebuff_OnTimeout()
    MyDebug "tmrNextRebuff Buff Buddy Timeout"

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
    
    'Call g_Objects.Equipment.Update
    
    Set m_colBuffs = New Collection
    
    With g_ui.Buffs
        
        'Buff Wand 1st
        If Valid(m_HeldItem) Then
            If (m_HeldItem.itemType = ITEM_WAND) Then
                CheckItemBuff .chkOtherItmWandHL.Checked, SPELL_HERMETIC_LINK, m_HeldItem
                CheckItemBuff .chkOtherItmWandDef.Checked, SPELL_DEF, m_HeldItem
                CheckItemBuff .chkOtherItmWandSD.Checked, SPELL_SPIRIT_DRINKER, m_HeldItem
            End If
        End If
        
        If .chkEnableOtherCreatureBuffs.Checked Then
            For i = 0 To .lstOtherCreature.Count - 1
                bChecked = .lstOtherCreature.Data(2, i, 0)
                sSpellFam = .lstOtherCreature.Data(1, i, 0)
                If bChecked Then
                    Call AddBuffToList(g_buffBuddy, sSpellFam)
                End If
            Next i
        End If
        
        If .chkEnableOtherLifeBuffs.Checked Then
            For i = 0 To .lstOtherLifePros.Count - 1
                bChecked = .lstOtherLifePros.Data(2, i, 0)
                sSpellFam = .lstOtherLifePros.Data(1, i, 0)
                If bChecked Then
                    Call AddBuffToList(g_buffBuddy, sSpellFam)
                End If
            Next i
        End If
        
        
        '*****************************************************************************
        'Special macro items which may not be currently equipped but need to be buffed
        '*****************************************************************************
        
        'Weapon / Bow / Shield
        If Valid(m_HeldShield) Then
            CheckItemBanes .chkOtherItmArmorShield.Checked, m_HeldShield
        End If
         
        If Valid(m_HeldItem) Then
            If (m_HeldItem.itemType = ITEM_MELEE_WEAPON) Then
                CheckItemBuff .chkOtherItmWeapHS.Checked, SPELL_HS, m_HeldItem
            End If
            If (m_HeldItem.itemType = ITEM_MELEE_WEAPON) Or (m_HeldItem.itemType = ITEM_MISSILE_WEAPON) Then
                CheckItemBuff .chkOtherItmWeapBD.Checked, SPELL_BD, m_HeldItem
                CheckItemBuff .chkOtherItmWeapSK.Checked, SPELL_SK, m_HeldItem
                CheckItemBuff .chkOtherItmWeapDE.Checked, SPELL_DEF, m_HeldItem
            End If
        End If

    End With

    MyDebug "BuildBuffList: total spells: " & m_colBuffs.Count

Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsBuddyBuffer.BuildBuffList - " & Err.Description & " line: " & Erl
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

    MyDebug "clsBuddyBuffer.StartService(" & iBuffMode & ", " & CStr(bRepeatCycles) & ", " & iContinuousBuffs & ")"
    
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
        MyDebug "clsBuddyBuffer.StartService - Pushing buffs to list"
        Call PushBuffs(False, True)
    End If
    
    'setup next rebuff time if required
    If m_bRepeat Then
        Call PrepareNextRebuffTime
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsBuddyBuffer.StartService"
    Resume Fin
End Sub

Public Sub StopService()
    MyDebug "clsBuddyBuffer.StopService()"
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
    
    'If m_tmrSecure.Expired Then
    '    PrintWarning "Buff Buddy Timeout - disabling!"
    '    MyDebug "clsBuddyBuffer.tmrSecure expired!"
    '    Call StopService
    '    Set g_buffBuddy = Nothing
    '    Call g_Macro.GoIdle
    '    Exit Sub
    'End If
    
    Set objBuff = m_BuffQueue.ReadTop
    
    'Item or Creature/Life buff ?
    If objBuff.SpellSchool = SCHOOL_ITEM Then
    
        'choose the right spell database
        Set spellList = g_Spells.Items
        
        'find the object we want to cast on
        Set objItem = g_Objects.FindObject(objBuff.TargetGUID)
        
        'Make sure we found an item to buff (should always be the case...)
        If Not Valid(objItem) Then
            PrintWarning "clsBuddyBuffer.CastNextSpell : couldn't find the target object to buff. Ignoring"
            Call m_BuffQueue.Pop    'remove the bad buff from queue
            GoTo Fin
        End If
        
        'Setup the casting params
        lTargetGUID = objItem.Guid
        sTargetName = objItem.Name
    
    Else    'Creature/Life buff
    
        'choose the right spell database
        Set spellList = g_Spells.oBuffs
        
        'find the object we want to cast on
        Set objItem = g_Objects.FindObject(objBuff.TargetGUID)
        
        'Make sure we found an item to buff (should always be the case...)
        If Not Valid(objItem) Then
            PrintWarning "clsBuddyBuffer.CastNextSpell : couldn't find the target object to buff. Ignoring"
            Call m_BuffQueue.Pop    'remove the bad buff from queue
            Call g_Macro.GoIdle     'Reset State
            GoTo Fin
        End If
        
        'Setup the casting params
        lTargetGUID = objItem.Guid
        sTargetName = objItem.Name
        m_sCurTargetName = sTargetName
        
    End If
    
    'MyDebug "Spell level wanted : " & objBuff.LevelWanted
    
    'Find the spell infos
    If objBuff.SpellType <> SPELLTYPE_NORMAL Then   'if it's a typed spell (i.e banes)
        Set objSpell = spellList.FindSpellByType(objBuff.SpellType, objBuff.SpellElement, g_Data.BuffsSpellsLevel)
        MyDebug "clsBuddyBuffer.NextSpell:Typed Spell: wanted: " & objBuff.SpellType
    Else
        Set objSpell = spellList.FindSpell(objBuff.SpellFamily, g_Data.BuffsSpellsLevel)
        MyDebug "clsBuddyBuffer.NextSpell: SpellFamily: " & objBuff.SpellFamily
    End If
    
    'Check if we have been able to find the spell, and cast it
    If Not Valid(objSpell) Then
        PrintErrorMessage "Could not find a valid spell for " & objBuff.Description & " - Igoring it."
        Call m_BuffQueue.Pop    'remove the bad buff from queue
        GoTo Fin
    Else
        m_sCurSpellName = objSpell.SpellName
        
        If Not (IsBusy) Then    ' Prevent spell spamming
        
            MyDebug "clsBuddyBuffer.CastNextSpell: Not (IsBusy), casting: " & objSpell.SpellName
            
            If Not (g_ui.Options.chkFilterLTMsg.Checked) Then
                If lTargetGUID = g_Objects.Player.Guid Or lTargetGUID = 0 Then
                    PrintMessage "[Rebuff] Casting " & objSpell.SpellName
                Else
                    PrintMessage "[Rebuff] Casting " & objSpell.SpellName & " on " & sTargetName
                End If
            End If
            
            Call g_Spells.CastThisSpell(objSpell, lTargetGUID)
        Else
            'MyDebug "clsBuddyBuffer.CastNextSpell: IsBusy, _not_ casting: " & objSpell.SpellName
        End If
    End If
    
    'Update secure timer
    Call m_tmrSecure.SetNextTime(SECURE_TIMER)
    m_tmrSecure.Enabled = True
    
Fin:
    Set objItem = Nothing
    Set objBuff = Nothing
    Set objSpell = Nothing
    Set spellList = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CastNextSpell"
    Resume Fin
End Sub

'Read the AC Console to figure out if we have casted the current buff spell,
'And pop it from the queue
'You cast Adja's Blessing on yourself
Public Sub CheckCastedSpell(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    If m_sCurSpellName <> "" Then
        If InStr(LCase(sCastMsg), LCase(m_sCurSpellName)) Then
            
            MyDebug "clsBuddyBuffer.CheckCastedSpell - Current spell " & m_sCurSpellName & " has been cast"
            
            If (g_Macro.State = ST_BUDDYREBUFF) Then
                Call m_BuffQueue.Pop
                Call g_Spells.c_SpellQueue.Pop
                'Update secure timer
                Call m_tmrSecure.Reset
                'update progress bar
                g_ui.Main.progBuffs.Value = g_ui.Main.progBuffs.Value + 1
                m_sCurSpellName = ""
            End If
            
            If g_Spells.Casting Then
                'MyDebug "clsBuddyBuffer.CheckCastedSpell - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        Else
            MyDebug "clsBuddyBuffer.CheckCastedSpell doesn't match: " & m_sCurSpellName & " : " & sCastMsg
            If g_Spells.Casting Then
                'MyDebug "clsBuddyBuffer.CheckCastedSpell - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        End If
    Else
        locDebug "clsBuddyBuffer.CheckCastedSpell: m_sCurSpellName is Blank"
        If g_Spells.Casting Then
            locDebug "clsBuddyBuffer.CheckCastedSpell - g_Spells.Casting is true, calling OnSpellCastComplete"
            Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
        End If

    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckCastedSpell - " & Err.Description
    Resume Fin
End Sub

'Read the AC Console to figure out if we are trying to cast on an Unenchantable item
Public Sub CheckUnenchantableItem(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    'm_sCurTargetName
    If m_sCurTargetName <> "" Then
        If InStr(LCase(sCastMsg), LCase(m_sCurTargetName)) Then
            
            locDebug "clsBuddyBuffer.CheckUnenchantableItem - Current spell " & m_sCurTargetName & " resisted our spell"
            
            If (g_Macro.State = ST_REBUFF) Then
                MyDebug "clsBuddyBuffer.CheckUnenchantableItem: poping off BuffQueue: " & m_sCurTargetName
                Call m_BuffQueue.Pop
                Call g_Spells.c_SpellQueue.Pop
                'update progress bar
                g_ui.Main.progBuffs.Value = g_ui.Main.progBuffs.Value + 1
                m_sCurTargetName = ""
            End If
            
            If g_Spells.Casting Then
                locDebug "clsBuddyBuffer.CheckUnenchantableItem - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        Else
            locDebug "clsBuddyBuffer.CheckUnenchantableItem doesn't match: " & m_sCurTargetName & " : " & sCastMsg
            If g_Spells.Casting Then
                locDebug "clsBuffer.CheckUnenchantableItem - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckUnechantableItem - " & Err.Description
    Resume Fin
End Sub



'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsBuddyBuffer] " & DebugMsg, bSilent)
    End If
End Sub
