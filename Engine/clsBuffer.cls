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

Private m_droppedQueue As clsSpellQueue     'list of dropped buff spells
Private m_buffItems As colObjects           'list of Items that need to be buffed
Private WithEvents m_tmrDroppedRebuff As clsTimer
Attribute m_tmrDroppedRebuff.VB_VarHelpID = -1

Private m_iRebuffMode As Integer
Private m_dRebuffInterval As Double         'time in seconds
Private m_dBuffCycleLen As Double           'time for a full buff cycle, in seconds
Private m_bRepeat As Boolean                'repeat buff cycles?

Private m_sCurSpellName As String           'Current spell name
Private m_sCurTargetName As String          'Buff Item Target name

Private WithEvents m_tmrNextRebuff As clsTimer
Attribute m_tmrNextRebuff.VB_VarHelpID = -1

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
    Set m_tmrDroppedRebuff = CreateTimer
    m_tmrDroppedRebuff.Enabled = False
    
    'default rebuff mode
    m_iRebuffMode = REBUFF_FULL
    m_dRebuffInterval = 10          'time in seconds
    m_iNumBuffs = 5
    m_bRepeat = True
    mNumBuffsToCast = 0
    
    Set m_droppedQueue = New clsSpellQueue
    Set m_buffItems = New colObjects
    
    Call Reset
End Sub

Private Sub Reset()
    Set m_BuffQueue = New clsSpellQueue
    Set m_colBuffs = New Collection
    Call m_tmrNextRebuff.Reset
    m_lCurBuff = -1
    m_sCurSpellName = ""
    m_sCurTargetName = ""
End Sub


Private Sub Class_Terminate()
    Set m_BuffQueue = Nothing
    Set m_colBuffs = Nothing
    Set m_buffItems = Nothing
    Set m_droppedQueue = Nothing
    Set m_tmrNextRebuff = Nothing
    Set m_tmrDroppedRebuff = Nothing
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
            .TargetGUID = objTarget.Guid
            .TargetName = objTarget.Name
        Else
            .TargetName = "Self"
            .TargetGUID = 0
        End If
        
    End With
    
    Call m_colBuffs.Add(objBuff, CStr(objBuff.Index))
    
Fin:
    Set objBuff = Nothing
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
    
    If Not Valid(objItem) Then Exit Sub
    If objItem.unEnchantable Then Exit Sub
    
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

Private Sub CheckAddBane(ByVal iElement As Integer, ByVal bChecked As Boolean, objItem As acObject, ByVal aFamily As String)
    If Not Valid(objItem) Then Exit Sub
    If objItem.unEnchantable Then Exit Sub
    
    If bChecked Then
        Call AddBuffToList(aFamily, SPELLTYPE_BANE, iElement, objItem, , SCHOOL_ITEM)
    End If
End Sub


Private Sub AddBanes(objItem As acObject)
On Error GoTo ErrorHandler

    If Not Valid(objItem) Then
        PrintErrorMessage "clsBuffer.AddBanes : invalid objItem"
        Exit Sub
    End If
    
    If objItem.unEnchantable Then
        MyDebug "clsBuffer.AddBanes: Can not enchant: " & objItem.Name
        Exit Sub
    End If
    
    With g_ui.Buffs
        Call CheckAddBane(DMG_NONE, .chkItmBaneImpen.Checked, objItem, SPELL_IMPEN)
        Call CheckAddBane(DMG_SLASHING, .chkItmBaneSlash.Checked, objItem, SPELL_BANE_SLASHING)
        Call CheckAddBane(DMG_PIERCING, .chkItmBanePierce.Checked, objItem, SPELL_BANE_PIERCING)
        Call CheckAddBane(DMG_BLUDGEONING, .chkItmBaneBludg.Checked, objItem, SPELL_BANE_BLUDEONING)
        Call CheckAddBane(DMG_FIRE, .chkItmBaneFire.Checked, objItem, SPELL_BANE_FIRE)
        Call CheckAddBane(DMG_COLD, .chkItmBaneFrost.Checked, objItem, SPELL_BANE_COLD)
        Call CheckAddBane(DMG_ACID, .chkItmBaneAcid.Checked, objItem, SPELL_BANE_ACID)
        Call CheckAddBane(DMG_LIGHTNING, .chkItmBaneLightning.Checked, objItem, SPELL_BANE_LIGHTNING)
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
    
    'See if we need to buff any weapons
    Call BuffWeapons
    
    m_lCurBuff = -1
End Sub

Private Sub FillBuffQueue(ByVal iFillMode As Integer, Optional ByVal bCleanQueue As Boolean = False)
On Error GoTo ErrorHandler

    Dim objBuff As clsSpellQueueItem
    
    If iFillMode = REBUFF_CONTINUOUS Then
        Dim i As Integer
        
        For i = 1 To m_iNumBuffs
            m_lCurBuff = m_lCurBuff + 1
            
            If m_lCurBuff > m_colBuffs.Count - 1 Then  ' count - 1 ? because we just added 1 to m_lCurBuff
                MyDebug "clsBuffer - Cont Buff Cycle over"
                Call OnRestartBuffCycle
                GoTo Fin
            End If
            
            If FindBuff(m_lCurBuff, objBuff) Then
                'Call addBuffToQueue(objBuff)
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
            'Call addBuffToQueue(objBuff)
            Call m_BuffQueue.Add(objBuff)
        Next objBuff
        
    End If
    
Fin:
    Set objBuff = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "AddBanes - " & Err.Description
    Resume Fin
End Sub

'Adds a buff to the Queue if it's not already there
Private Sub addBuffToQueue(aBuff As clsSpellQueueItem)
On Error GoTo ErrorHandler
    Dim objBuff As clsSpellQueueItem
    Dim i As Integer

    For Each objBuff In m_BuffQueue.SpellQueue
        If (objBuff.SpellFamily = aBuff.SpellFamily) And (objBuff.TargetGUID = 0) Then
            'Already on the list, so exit sub
            MyDebug "clsBuffer: Already in BuffQueue: " & aBuff.Description
            Exit Sub
        End If
    Next objBuff
    
    MyDebug "clsBuffer: Adding to BuffQueue: " & aBuff.Description
    
    'Not found on Queue, so add it
    Call m_BuffQueue.Add(aBuff)
    
Fin:
    Set objBuff = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "addBuffToQueue - " & Err.Description
    Resume Fin
End Sub


'Adds a buff to the Dropped Buff Queue if it's not already there
Private Sub addDroppedBuffToQueue(aBuff As clsSpellQueueItem)
On Error GoTo ErrorHandler
    Dim objBuff As clsSpellQueueItem
    Dim i As Integer

    For Each objBuff In m_droppedQueue.SpellQueue
        If objBuff.SpellFamily = aBuff.SpellFamily Then
            'Already on the list, so exit sub
            locDebug "clsBuffer.addDropped: In m_droppedQueue: " & aBuff.Description
            Exit Sub
        End If
    Next objBuff
    
    MyDebug "clsBuffer.addDropped: Add to m_droppedQueue: " & aBuff.Description
    
    'Not found on Queue, so add it
    Call m_droppedQueue.Add(aBuff)
    
    Call m_tmrDroppedRebuff.SetNextTime(3)
    m_tmrDroppedRebuff.Enabled = True
    
Fin:
    Set objBuff = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "addDroppedBuffToQueue - " & Err.Description
    Resume Fin
End Sub

Private Sub PrepareNextRebuffTime()
    'setup timer
    Call m_tmrNextRebuff.SetNextTime(m_dRebuffInterval)
    ' See if we have any mana stones to charge
    g_Macro.Loot.idManaStones
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

Private Sub m_tmrDroppedRebuff_OnTimeout()
On Error GoTo ErrorHandler
    
    Dim objBuff As clsSpellQueueItem
    Dim buffSpell As clsSpell
    Dim objItem As acObject
    Dim testItem As WorldObject
    Dim bMissingSpell As Boolean
    Dim bFoundSpell As Boolean
    Dim iCount As Integer
    
    MyDebug "tmrDroppedRebuff OnTimeout: m_buffItems.count: " & m_buffItems.Count & " :dQ: " & m_droppedQueue.Count
    
    bMissingSpell = False
    iCount = 0

    'List of dropped Item buffs
3    For Each objBuff In m_droppedQueue.SpellQueue
        'Set buffSpell = g_Spells.Items.FindSpellByType(objBuff.SpellType, objBuff.SpellElement, objBuff.LevelWanted)
4        Set buffSpell = g_Spells.Items.FindSpell(objBuff.SpellFamily)
        If Valid(buffSpell) Then
            'Go through all the Equipped Items and see if this spell is missing from them
            MyDebug "tmrDroppedRebuff: buffSpell: " & buffSpell.SpellName

5            For Each objItem In m_buffItems
                bMissingSpell = False
                bFoundSpell = False
                
                'Set testItem = g_Filters.g_worldFilter.Item(objItem.Guid)
                'If Valid(testItem) Then
                '    If testItem.Longs(keyActiveSpellCount) > 0 Then
                '        Dim spellCount As Integer
                '        Dim i As Integer
                '        spellCount = testItem.Longs(keyActiveSpellCount)
                '        MyDebug "tmrDroppedRebuff: ActiveSpellCount:  " & spellCount
                '        For i = 0 To spellCount - 1
                '            If testItem.ActiveSpell(i) = buffSpell.SpellID Then
                '                'SpellID = wObj.spell(i)
                '                MyDebug "m_tmrDroppedRebuff: ** TEST Buff Exists **: " & objBuff.Description & " : " & testItem.Name
                '                bMissingSpell = False
                '            End If
                '        Next i
                '    Else
                '        'No Active spells, so must be missing them
                '        MyDebug "m_tmrDroppedRebuff: ** TEST Missing Buff **: " & objBuff.Description & " : " & testItem.Name
                '        bMissingSpell = True
                '        iCount = iCount + 1
                '    End If
                'End If
                'If bMissingSpell Then
                '    iCount = iCount + 1
                'End If
                
                If Valid(objItem) Then
6                    If objItem.Spells.Exists(buffSpell.SpellName) Then
                        MyDebug "m_tmrDroppedRebuff: ** Buff Exists **: " & objBuff.Description & " : " & objItem.Name
                        'bMissingSpell = False
                    Else
                        MyDebug "m_tmrDroppedRebuff: Missing Buff: " & objBuff.Description & " : " & objItem.Name
                        bMissingSpell = True
                        iCount = iCount + 1
                    End If
                End If
7            Next objItem
            
            If iCount >= (m_buffItems.Count / 2) Then
                MyDebug "m_tmrDroppedRebuff: Queing objBuff: " & objBuff.Description & " :: " & objBuff.TargetName
                Call addBuffToQueue(objBuff)
            Else
                MyDebug "m_tmrDroppedRebuff: Only missing (" & iCount & "): " & buffSpell.SpellName
            End If
            
            'If bMissingSpell Then
            '    MyDebug "m_tmrDroppedRebuff: Queing objBuff: " & objBuff.Description & " :: " & objBuff.TargetName
            '    Call addBuffToQueue(objBuff)
            'Else
            '    MyDebug "m_tmrDroppedRebuff: no missing spells found for: " & buffSpell.SpellName
            'End If
            'If bFoundSpell Then
            '    MyDebug "m_tmrDroppedRebuff: Found an active spell: " & buffSpell.SpellName
            'Else
            '    MyDebug "m_tmrDroppedRebuff: Queing objBuff: " & objBuff.Description & " :: " & objBuff.TargetName
            '    Call addBuffToQueue(objBuff)
            'End If
        End If
    Next objBuff
    
    m_tmrDroppedRebuff.Enabled = False

Fin:
    Set m_droppedQueue = Nothing
    Set m_buffItems = Nothing
    Set objBuff = Nothing
    Set objItem = Nothing
    Set buffSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "m_tmrDroppedRebuff_OnTimeout - " & Err.Description & " - line: " & Erl
    Resume Fin
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
        Call CheckItemBuff(.chkItmWandHL.Checked, SPELL_HERMETIC_LINK, g_Data.Wand)
        Call CheckItemBuff(.chkItmWandDef.Checked, SPELL_DEF, g_Data.Wand)
        Call CheckItemBuff(.chkItmWandSD.Checked, SPELL_SPIRIT_DRINKER, g_Data.Wand)
        
        'Weapon / Bow / Shield
        If IsMelee Then
            Dim objWeapon As acObject
            
            If g_Macro.CombatType = TYPE_MELEE Then
                Set objWeapon = g_Data.Weapon
                Call CheckItemBanes(.chkItmArmorShield.Checked, g_Data.Shield)
                Call CheckItemBuff(.chkItmWeapHS.Checked, SPELL_HS, objWeapon)
            Else    'Archer
                Set objWeapon = g_Data.Bow
            End If
            
            Call CheckItemBuff(.chkItmWeapBD.Checked, SPELL_BD, objWeapon)
            Call CheckItemBuff(.chkItmWeapSK.Checked, SPELL_SK, objWeapon)
            Call CheckItemBuff(.chkItmWeapDE.Checked, SPELL_DEF, objWeapon)
        End If
        
        'Armor Banes
        If .chkItmSelectSelf.Checked Then
            Call CheckItemBanes(.chkItmSelectSelf.Checked, g_Objects.Player)
        Else
            Call CheckItemBanes(.chkItmArmorHead.Checked, g_Objects.Equipment.head)
            Call CheckItemBanes(.chkItmArmorHands.Checked, g_Objects.Equipment.Hands)
            Call CheckItemBanes(.chkItmArmorFeet.Checked, g_Objects.Equipment.Feet)
            Call CheckItemBanes(.chkItmArmorTopUndie.Checked, g_Objects.Equipment.TopUndie)
            Call CheckItemBanes(.chkItmArmorBottomUndie.Checked, g_Objects.Equipment.BottomUndie)
            Call CheckCollectionBanes(.chkItmArmorTop.Checked, g_Objects.Equipment.Top)
            Call CheckCollectionBanes(.chkItmArmorBottom.Checked, g_Objects.Equipment.Bottom)
        End If
        
        'Critter Buffs
        If .chkEnableCreatureBuffs.Checked Then
            For i = 0 To .lstCreature.Count - 1
                bChecked = .lstCreature.Data(2, i, 0)
                sSpellFam = .lstCreature.Data(1, i, 0)
                If bChecked Then
                    Call AddBuffToList(sSpellFam)
                End If
            Next i
        End If
        
        'Life Pro's
        If .chkEnableLifeBuffs.Checked Then
            For i = 0 To .lstLifePros.Count - 1
                bChecked = .lstLifePros.Data(2, i, 0)
                sSpellFam = .lstLifePros.Data(1, i, 0)
                If bChecked Then
                    Call AddBuffToList(sSpellFam)
                End If
            Next i
        End If

    End With

    MyDebug "BuildBuffList: total spells: " & m_colBuffs.Count
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "BuildBuffList - " & Err.Description
    Resume Fin
End Sub

'Walks Buff List and makes sure that spell is currently active on this toon
Private Sub existingSpellList()
On Error GoTo ErrorHandler
    
    Dim objBuff As New clsSpellQueueItem
    Dim objSpell As clsSpell
    Dim isFound As Boolean
    Dim i As Long

    locDebug "existingSpellList: Existing Spells: " & g_Filters.ActiveSpellsCount

    'copy all the spells from the buff list to the pending buffs queue
    For Each objBuff In m_colBuffs

        If objBuff.SpellType <> SPELLTYPE_NORMAL Then
            'Do nothing
        Else
            isFound = False
            Set objSpell = g_Spells.sBuffs.FindSpell(objBuff.SpellFamily, g_Data.BuffsSpellsLevel)
            If Valid(objSpell) Then
                For i = 0 To (g_Filters.ActiveSpellsCount - 1)
                    'MyDebug "ID: " & g_Filters.ActiveSpell(i).SpellID & " :F: " & g_Filters.ActiveSpell(i).Family & " :T: " & g_Filters.ActiveSpell(i).TimeRemaining
                    If g_Filters.ActiveSpell(i).SpellID = objSpell.SpellID Then
                        isFound = True
                        GoTo found
                    End If
                Next i
            Else
                isFound = True
            End If
found:
            If Not isFound Then
                 locDebug "existingSpellList: Missing spell, add BuffQueue: " & objBuff.Description
                 Call addBuffToQueue(objBuff)
            End If
        End If
    
    Next objBuff

Fin:
    Set objBuff = Nothing
    Set objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "existingSpellList - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

'See if this Family of spell is active on the player
Public Function existingSpellFamilyOnPlayer(ByVal aFamily As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim objSpell As clsSpell
    Dim isFound As Boolean
    Dim i As Long

    isFound = False
    
    For i = 0 To (g_Filters.ActiveSpellsCount - 1)
    
        MyDebug "existingSpellFamily: " & g_Filters.ActiveSpell(i).Family & " :vs: " & aFamily
        
        Set objSpell = g_Spells.sBuffs.FindSpell(aFamily, g_Data.BuffsSpellsLevel)

        If Valid(objSpell) Then
            'MyDebug "ID: " & g_Filters.ActiveSpell(i).SpellID & " :F: " & g_Filters.ActiveSpell(i).Family & " :T: " & g_Filters.ActiveSpell(i).TimeRemaining
            If g_Filters.ActiveSpell(i).SpellID = objSpell.SpellID Then
                isFound = True
                GoTo Fin
            End If
        End If
    Next i

Fin:
    existingSpellFamilyOnPlayer = isFound
    Set objSpell = Nothing
    Exit Function
ErrorHandler:
    PrintErrorMessage "existingSpellFamilyOnPlayer - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function existingSpellsOnPlayer(ByVal aFamily As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim newSpell As New clsSpellQueueItem
    Dim oldSpell As New clsSpellQueueItem
    Dim isFound As Boolean
    Dim i As Long
    Dim iSpell As Long

    locDebug "Displaying Existing Spells: " & g_Filters.ActiveSpellsCount

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
    existingSpellsOnPlayer = isFound
    Set newSpell = Nothing
    Set oldSpell = Nothing
    Exit Function
ErrorHandler:
    PrintErrorMessage "existingSpellsOnPlayer - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

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

    'See if we need any "pre-buffs"
    Call PreBuffCheck
    
    'See if we need to buff up Any weapons
    Call BuffWeapons
    
    'push buffs to the queue if required
    If bPushBuffsNow Then
        MyDebug "StartService - Pushing buffs to list"
        Call PushBuffs(False, True)
        'Make sure we also cast any spells that are missing from the Player
        Call existingSpellList
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

'Check to see if we need to pre-buff certain spells
Private Sub PreBuffCheck()
On Error GoTo ErrorHandler
    Dim objBuff As clsSpellQueueItem
    Dim aFamily As String

    'First check critter buffs
    If g_ui.Buffs.chkEnableCreatureBuffs.Checked Then
        'First check critter
        aFamily = "Creature Enchantment Mastery Self"
        If Not existingSpellFamilyOnPlayer(aFamily) Then
            Set objBuff = New clsSpellQueueItem
            objBuff.SpellFamily = aFamily
            objBuff.SpellType = SPELLTYPE_NORMAL
            objBuff.TargetName = "Self"
            objBuff.TargetGUID = 0
            objBuff.SpellSchool = SCHOOL_CREATURE
            Call addBuffToQueue(objBuff)
        End If
        'Then check Will
        aFamily = "Willpower Self"
        If Not existingSpellFamilyOnPlayer(aFamily) Then
            Set objBuff = New clsSpellQueueItem
            objBuff.SpellFamily = aFamily
            objBuff.SpellType = SPELLTYPE_NORMAL
            objBuff.TargetName = "Self"
            objBuff.TargetGUID = 0
            objBuff.SpellSchool = SCHOOL_CREATURE
            Call addBuffToQueue(objBuff)
        End If
        'Then check Focus
        aFamily = "Focus Self"
        If Not existingSpellFamilyOnPlayer(aFamily) Then
            Set objBuff = New clsSpellQueueItem
            objBuff.SpellFamily = aFamily
            objBuff.SpellType = SPELLTYPE_NORMAL
            objBuff.TargetName = "Self"
            objBuff.TargetGUID = 0
            objBuff.SpellSchool = SCHOOL_CREATURE
            Call addBuffToQueue(objBuff)
        End If
        'Then check Life
        aFamily = "Life Magic Mastery Self"
        If Not existingSpellFamilyOnPlayer(aFamily) Then
            Set objBuff = New clsSpellQueueItem
            objBuff.SpellFamily = aFamily
            objBuff.SpellType = SPELLTYPE_NORMAL
            objBuff.TargetName = "Self"
            objBuff.TargetGUID = 0
            objBuff.SpellSchool = SCHOOL_CREATURE
            Call addBuffToQueue(objBuff)
        End If
        'And last check Item
        aFamily = "Item Enchantment Mastery Self"
        If Not existingSpellFamilyOnPlayer(aFamily) Then
            Set objBuff = New clsSpellQueueItem
            objBuff.SpellFamily = aFamily
            objBuff.SpellType = SPELLTYPE_NORMAL
            objBuff.TargetName = "Self"
            objBuff.TargetGUID = 0
            objBuff.SpellSchool = SCHOOL_CREATURE
            Call addBuffToQueue(objBuff)
        End If
    End If
    
Fin:
    Set objBuff = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PreBuffCheck - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub buffItem(ByVal isChecked As Boolean, ByVal sFamily As String, ByVal objItem As acObject)
On Error GoTo ErrorHandler
    Dim objBuff As New clsSpellQueueItem
    
    If isChecked Then
        objBuff.SpellFamily = sFamily
        objBuff.SpellType = SPELLTYPE_NORMAL
        objBuff.LevelWanted = 8
        objBuff.TargetName = objItem.Name
        objBuff.TargetGUID = objItem.Guid
        objBuff.SpellSchool = SCHOOL_ITEM
        Call addBuffToQueue(objBuff)
    End If

Fin:
    Set objBuff = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "buffItem - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

' Buff all the extra weapons
Public Sub BuffWeapons()
On Error GoTo ErrorHandler
    Dim aWeap As acObject
    Dim i As Integer
    
    If Not g_ui.Macro.chkAutoSwitch.Checked Then
        Exit Sub
    End If
    
    'Buff extra Weapons/Wands
    For i = DMG_SLASHING To DMG_LIGHTNING
    
        Set aWeap = Nothing
        
        Select Case i
            Case DMG_SLASHING
                If g_ui.Macro.chkWeapSlash.Checked Then
                    Set aWeap = g_Data.WeapSlash
                End If
            Case DMG_BLUDGEONING
                If g_ui.Macro.chkWeapBludge.Checked Then
                    Set aWeap = g_Data.WeapBludge
                End If
            Case DMG_PIERCING
                If g_ui.Macro.chkWeapPierce.Checked Then
                    Set aWeap = g_Data.WeapPierce
                End If
            Case DMG_FIRE
                If g_ui.Macro.chkWeapFire.Checked Then
                    Set aWeap = g_Data.WeapFire
                End If
            Case DMG_COLD
                If g_ui.Macro.chkWeapCold.Checked Then
                    Set aWeap = g_Data.WeapCold
                End If
            Case DMG_ACID
                If g_ui.Macro.chkWeapAcid.Checked Then
                    Set aWeap = g_Data.WeapAcid
                End If
            Case DMG_LIGHTNING
                If g_ui.Macro.chkWeapLight.Checked Then
                    Set aWeap = g_Data.WeapLight
                End If
        End Select
        
        If Valid(aWeap) Then
            If g_Macro.CombatType = TYPE_CASTER Then
                Call buffItem(g_ui.Buffs.chkItmWandHL.Checked, SPELL_HERMETIC_LINK, aWeap)
                Call buffItem(g_ui.Buffs.chkItmWandDef.Checked, SPELL_DEF, aWeap)
                Call buffItem(g_ui.Buffs.chkItmWandSD.Checked, SPELL_SPIRIT_DRINKER, aWeap)
            ElseIf g_Macro.CombatType = TYPE_MELEE Then
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapHS.Checked, SPELL_HS, aWeap)
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapBD.Checked, SPELL_BD, aWeap)
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapSK.Checked, SPELL_SK, aWeap)
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapDE.Checked, SPELL_DEF, aWeap)
            Else    'Archer
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapBD.Checked, SPELL_BD, aWeap)
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapSK.Checked, SPELL_SK, aWeap)
                Call CheckItemBuff(g_ui.Buffs.chkItmWeapDE.Checked, SPELL_DEF, aWeap)
            End If
        End If
        
    Next i

Fin:
    Set aWeap = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "WeaponBuff - " & Err.Description & " - line: " & Erl
    Resume Fin
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
        m_sCurTargetName = sTargetName
    
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
    PrintErrorMessage "CastNextSpell - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

'Receives text from Events and figures out if we need to recast an Item buff
Public Sub CheckExpiredItemSpell(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    Dim sSpellName As String
    Dim sTargetName As String
    Dim objBuff As clsSpellQueueItem
    Dim objSpell As clsSpell
    Dim buffSpell As clsSpell
    Dim objItem As acObject
    Dim bFound As Boolean
    Dim i As Long
    
'   The spell Cragstone's Will on Opal Blunt Baton has expired.
'   Adja's Blessing has Expired
    
    Dim iLen As Integer, aPos As Integer, bPos As Integer
    Dim idString As String, aString As String, bString As String
    
    idString = Replace(sCastMsg, "The spell ", "")
    idString = Replace(idString, " has expired.", "")
'   Cragstone's Will on Opal Blunt Baton
    locDebug "CheckExpiredItemSpell: idString: " & idString
    aPos = InStr(idString, " on ")
    iLen = Len(idString)
    If (aPos <= 0) Then aPos = 1
    sSpellName = Left(idString, aPos - 1)
    sTargetName = Right(idString, (iLen - aPos - 3))
    
    locDebug "CheckExpiredItemSpell: sSpellName: " & sSpellName & " :T: " & sTargetName
    
    bFound = False
    
    Set objSpell = g_Spells.Items.FindSpellByName(sSpellName)
1
    If Not Valid(objSpell) Then
        ' Can't find this spell anywhere
        locDebug "CheckExpiredItemSpell: not a valid Item spell: " & sSpellName
        Exit Sub
    End If
2
    'Find the item and see if it still has that spell active on it
3
    Set objItem = g_Objects.Items.FindByName(sTargetName, True)
    If Valid(objItem) Then
        If (objItem.Name = sTargetName) And objItem.Equiped Then
4
            If Not m_buffItems.Exists(objItem.Guid) Then
                Call m_buffItems.addObject(objItem)
            End If
            
            Call g_Macro.Loot.SetSilentID(True)
            Call g_Hooks.IDQueueAdd(objItem.Guid)
            
            bFound = True
            
            'If objItem.Spells.Count > 0 And objItem.Spells.Exists(sSpellName) Then
5           '     ' Found it!
            '    MyDebug "CheckExpiredItemSpell: Found spell on item: " & sSpellName
            '    Exit Sub
            'End If
        End If
    End If
6
    If Not bFound Then
        MyDebug "CheckExpiredItemSpell: Could not find item: " & sTargetName
        Exit Sub
    End If

    'MyDebug "CheckExpiredItemSpell:  objSpell: " & objSpell.SpellName & " : " & objSpell.SpellFamily

    'If not then we need to see if it's a buff we would cast (in the buff queue/list)
    For Each objBuff In m_colBuffs
        If objBuff.SpellSchool = SCHOOL_ITEM Then   'if it's an Item spell (i.e banes)
            'Set buffSpell = g_Spells.Items.FindSpellByType(objBuff.SpellType, objBuff.SpellElement, g_Data.BuffsSpellsLevel)
            If Valid(objSpell) Then
                If (objBuff.SpellFamily = objSpell.SpellFamily) Then
                    MyDebug "CheckExpiredItemSpell: Queing objBuff: " & objBuff.Description & " :: " & objBuff.TargetName
                    'Call addBuffToQueue(objBuff)
                    Call addDroppedBuffToQueue(objBuff)
                    Exit Sub
                End If
            End If
        End If
    Next objBuff

Fin:
    Set objItem = Nothing
    Set objBuff = Nothing
    Set objSpell = Nothing
    Set buffSpell = Nothing
    Exit Sub
ErrorHandler:
    'PrintErrorMessage "CheckExpiredItemSpell - " & Err.Description & " - line: " & Erl
    MyDebug "CheckExpiredItemSpell: ERROR: - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

'Look for "has expired" messages to make sure we don't drop any buffs
Public Sub CheckExpiredSpell(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    Dim sSpellName As String
    Dim objBuff As clsSpellQueueItem
    Dim objSpell As clsSpell
    Dim buffSpell As clsSpell
    Dim i As Long
    
'   The spell Cragstone's Will on Opal Blunt Baton has expired.
'   Adja's Blessing has Expired
    
    sSpellName = Replace(sCastMsg, " has expired.", "")
    locDebug "CheckExpiredSpell: " & sSpellName
    
    Set objSpell = g_Spells.sBuffs.FindSpellByName(sSpellName)
    
    If Not Valid(objSpell) Then
        ' Can't find this spell anywhere
        MyDebug "CheckExpiredSpell: not a valid spell name: " & sSpellName
        Exit Sub
    End If

    'First, check to see if we have this spell still active on us
    For i = 0 To (g_Filters.ActiveSpellsCount - 1)
        If g_Filters.ActiveSpell(i).SpellID = objSpell.SpellID _
         And g_Filters.ActiveSpell(i).TimeRemaining > 200 Then
            ' Yep, found active spell, so we are A-OK!
            locDebug "CheckExpiredSpell: found active spell: " & g_Filters.ActiveSpell(i).TimeRemaining
            Exit Sub
        End If
    Next i
    
    'If not then we need to see if it's a buff we would cast (in the buff queue/list)
    'For i = 1 To m_iNumBuffs
    '    If FindBuff(i, objBuff) Then
    '        If objBuff.SpellID = objSpell.SpellID Then
    '            Call addBuffToQueue(objBuff)
    '            MyDebug "CheckExpiredSpell: found in m_colBuffs, adding to current Queue"
    '            Exit Sub
    '        End If
    '    End If
    'Next i
    
    'If not then we need to see if it's a buff we would cast (in the buff queue/list)
    For Each objBuff In m_colBuffs
        If objBuff.SpellSchool <> SCHOOL_ITEM Then
            Set buffSpell = g_Spells.sBuffs.FindSpell(objBuff.SpellFamily, g_Data.BuffsSpellsLevel)
            If Valid(buffSpell) Then
                If buffSpell.SpellID = objSpell.SpellID Then
                    MyDebug "CheckExpiredSpell: adding to current Queue: " & objBuff.Description
                    Call addBuffToQueue(objBuff)
                    Exit Sub
                End If
            End If
        End If
    Next objBuff
    
Fin:
    Set objBuff = Nothing
    Set objSpell = Nothing
    Set buffSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "CheckExpiredSpell - " & Err.Description
    Resume Fin
End Sub
        
        
'Read the AC Console to figure out if we have cast the current buff spell, if so pop it from the queue
'ex: You cast Adja's Blessing on yourself
Public Sub CheckCastedSpell(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    If m_sCurSpellName <> "" Then
        If InStr(LCase(sCastMsg), LCase(m_sCurSpellName)) Then
            
            locDebug "clsBuffer.CheckCastedSpell - Current spell " & m_sCurSpellName & " has been cast"
            
            If (g_Macro.State = ST_REBUFF) Then
                MyDebug "clsBuffer.CheckCastedSpell: poping off BuffQueue: " & m_sCurSpellName
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

'Read the AC Console to figure out if we are trying to cast on an Unenchantable item
Public Sub CheckUnenchantableItem(ByVal sCastMsg As String)
On Error GoTo ErrorHandler

    'm_sCurTargetName
    If m_sCurTargetName <> "" Then
        If InStr(LCase(sCastMsg), LCase(m_sCurTargetName)) Then
            
            locDebug "clsBuffer.CheckCastedSpell - Current spell " & m_sCurTargetName & " resisted our spell"
            
            If (g_Macro.State = ST_REBUFF) Then
                MyDebug "clsBuffer.CheckUnenchantableItem: poping off BuffQueue: " & m_sCurTargetName
                Call m_BuffQueue.Pop
                Call g_Spells.c_SpellQueue.Pop
                'update progress bar
                g_ui.Main.progBuffs.Value = g_ui.Main.progBuffs.Value + 1
                m_sCurTargetName = ""
            End If
            
            If g_Spells.Casting Then
                locDebug "clsBuffer.CheckUnenchantableItem - g_Spells.Casting is true, calling OnSpellCastComplete"
                Call g_Spells.OnSpellCastComplete(True, "CheckCastedSpell")
            End If
        Else
            locDebug "clsBuffer.CheckUnenchantableItem doesn't match: " & m_sCurTargetName & " : " & sCastMsg
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
        Call MyDebug("[clsBuffer] " & DebugMsg, bSilent)
    End If
End Sub


