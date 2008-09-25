VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Macro Spellcasting Class                          [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False

'Deals with the game's magic mode
Private Const MODULE_DESCRIPTION = "Spells"
Private Const SECURE_CASTING_TIME = 4

Public c_SpellQueue As clsSpellQueue
Public c_PrioritySpellQueue As clsSpellQueue

Private m_bCasting As Boolean
Private m_bReadyToCast As Boolean       'in turbo mode, bReadyToCast will move back to true after m_bCasting resets
                                        'this is used to tell if the spell cast is complete (EV_READY), even
                                        'when the Cast Breaker is running
                                        
Private m_bSpellsLoaded As Boolean

Private m_Vitals As clsSpellList       'MacroSpells.cvs (stam2health, stam2mana, etc)
Private m_sBuffs As clsSpellList       'BuffsSelf.cvs
Private m_oBuffs As clsSpellList       'BuffsOther.cvs
Private m_Items As clsSpellList        'ItemSpells.cvs
Private m_War As clsSpellList          'WarSpells.cvs
Private m_Debuffs As clsSpellList      'Debuffs.cvs
    
Private m_CurrentWarSpell As clsSpell
Private m_CurrentVuln As clsSpell
Private m_CurrentDamageType As eDamageType
    
Private m_DefaultDamageType As eDamageType
Private m_DefaultWarBolt As clsSpell
Private m_DefaultVuln As clsSpell

Private m_tmrPrecastDelay As clsTimer
Private WithEvents m_tmrSecureCasting As clsTimer
Attribute m_tmrSecureCasting.VB_VarHelpID = -1
Private WithEvents m_tmrCastBreaker As clsTimer
Attribute m_tmrCastBreaker.VB_VarHelpID = -1
Private m_dLastCastAttempt As Double

Private m_dicSpellNames As Dictionary   'dictionary of of the spell names, by spell id
                                        'used to quickly figure out the name of a spell from its spell ID

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    
    m_bSpellsLoaded = False
    
    m_DefaultDamageType = DMG_PIERCING
    
    Set m_tmrPrecastDelay = CreateTimer
    Set m_tmrSecureCasting = CreateTimer
    Set m_tmrCastBreaker = CreateTimer
    
    Set m_dicSpellNames = New Dictionary
    
    Set c_SpellQueue = New clsSpellQueue
    Set c_PrioritySpellQueue = New clsSpellQueue

    Call Reset
    
End Sub

Public Sub Reset()
    
    Set m_Vitals = New clsSpellList
    Set m_sBuffs = New clsSpellList
    Set m_oBuffs = New clsSpellList
    Set m_Items = New clsSpellList
    Set m_War = New clsSpellList
    Set m_Debuffs = New clsSpellList
    
    m_Debuffs.Description = "Debuffs"
    m_Items.Description = "Items"
    m_oBuffs.Description = "oBuffs"
    m_sBuffs.Description = "sBuffs"
    m_Vitals.Description = "Vitals"
    m_War.Description = "War"
            
    Set m_CurrentVuln = New clsSpell
    Set m_CurrentWarSpell = New clsSpell
    
    Set m_DefaultWarBolt = New clsSpell
    Set m_DefaultVuln = New clsSpell
    'm_DefaultDamageType = DMG_PIERCING
    
    Call m_tmrPrecastDelay.Reset
    Call m_tmrSecureCasting.Reset
    Call m_tmrCastBreaker.Reset
    
    m_dLastCastAttempt = 0
    m_bCasting = False

End Sub

Private Sub Class_Terminate()
    Set c_SpellQueue = Nothing
    Set c_PrioritySpellQueue = Nothing
    
    Set m_Vitals = Nothing
    Set m_sBuffs = Nothing
    Set m_oBuffs = Nothing
    Set m_Items = Nothing
    Set m_War = Nothing
    Set m_Debuffs = Nothing
    
    Set m_CurrentVuln = Nothing
    Set m_CurrentWarSpell = Nothing
    
    Set m_DefaultWarBolt = Nothing
    Set m_DefaultVuln = Nothing
    
    Set m_tmrPrecastDelay = Nothing
    Set m_tmrCastBreaker = Nothing
    Set m_tmrSecureCasting = Nothing
        
    Set m_dicSpellNames = Nothing
    
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get Casting() As Boolean
    Casting = m_bCasting
End Property

Public Property Get Vitals() As clsSpellList
    Set Vitals = m_Vitals
End Property

Public Property Get sBuffs() As clsSpellList
    Set sBuffs = m_sBuffs
End Property

Public Property Get oBuffs() As clsSpellList
    Set oBuffs = m_oBuffs
End Property

Public Property Get Items() As clsSpellList
    Set Items = m_Items
End Property

Public Property Get War() As clsSpellList
    Set War = m_War
End Property

Public Property Get Debuffs() As clsSpellList
    Set Debuffs = m_Debuffs
End Property

Public Property Get CurrentVuln() As clsSpell
    Set CurrentVuln = m_CurrentVuln
End Property

Public Property Get CurrentWarSpell() As clsSpell
    Set CurrentWarSpell = m_CurrentWarSpell
End Property

Public Property Get CurrentDamage() As Integer
    CurrentDamage = m_CurrentDamageType
End Property

Public Property Get DefaultWarBolt() As clsSpell
    Set DefaultWarBolt = m_DefaultWarBolt
End Property

Public Property Get DefaultVuln() As clsSpell
    Set DefaultVuln = m_DefaultVuln
End Property

Public Property Get DefaultDamage() As Integer
    DefaultDamage = m_DefaultDamageType
End Property

Public Property Get tmrPrecastDelay() As clsTimer
    Set tmrPrecastDelay = m_tmrPrecastDelay
End Property

Public Property Get LastCastAttempt() As Double
    LastCastAttempt = m_dLastCastAttempt
End Property

'Public Property Get ReadyToCast() As Boolean
'    ReadyToCast = m_bReadyToCast
'End Property

'Public Property Let ReadyToCast(ByVal bVal As Boolean)
'    m_bReadyToCast = bVal
'End Property

'Public Property Get SpellNamesDic() As Dictionary
'    Set SpellNamesDic = m_dicSpellNames
'End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   Public Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function LoadSpells() As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    
    MyDebug "Loading Spells Databases"
    
    bRet = Not m_Vitals.LoadData(PATH_SPELLS & "\" & FILE_VITALSPELLS)
    If (Not bRet) Then bRet = Not m_War.LoadData(PATH_SPELLS & "\" & FILE_WARSPELLS)
    If (Not bRet) Then bRet = Not m_sBuffs.LoadData(PATH_SPELLS & "\" & FILE_SELFBUFFS)
    If (Not bRet) Then bRet = Not m_oBuffs.LoadData(PATH_SPELLS & "\" & FILE_OTHERBUFFS)
    If (Not bRet) Then bRet = Not m_Debuffs.LoadData(PATH_SPELLS & "\" & FILE_DEBUFFSPELLS)
    If (Not bRet) Then bRet = Not m_Items.LoadData(PATH_SPELLS & "\" & FILE_ITEMSPELLS)
    'If (Not bRet) Then bRet = Not LoadSpellNames(PATH_SPELLS & "\" & FILE_SPELLNAMES)
    
    m_bSpellsLoaded = Not bRet
    LoadSpells = m_bSpellsLoaded
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.LoadSpells - " & Err.Description
    Resume Fin
End Function


Public Sub OnSpellCastBegin(Optional ByVal aSource As String)
On Error GoTo ErrorHandler

    'set casting flags
    m_bCasting = True
    m_bReadyToCast = False
    Call m_tmrSecureCasting.SetNextTime(SECURE_CASTING_TIME)

    locDebug "clsSpells.OnSpellCastBegin from: " & aSource
    
    'If TurboMode Then
    '    locDebug "-- Starting Cast Breaker Timer :: " & g_Data.CastBreakerTime
    '    Call m_tmrCastBreaker.SetNextTime(g_Data.CastBreakerTime)
    'End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.OnSpellCastBegin - " & Err.Description
    Resume Fin
End Sub

Public Sub OnSpellCastComplete(Optional ByVal bSuccess As Boolean = True, Optional ByVal sSource As String = "")
On Error GoTo ErrorHandler

    locDebug "clsSpells.OnSpellCastComplete : " & CStr(bSuccess) & " - " & sSource
    
    ' See if this spell is on the Queue, if so, pop it off
    If bSuccess And (g_Spells.c_SpellQueue.Count > 0) Then
        locDebug "OnSpellCastBegin: c_SpellQueue.Count > 0"
        Call g_Spells.c_SpellQueue.Pop ' Pop the spell off the queue
    End If
    
    'c_PrioritySpellQueue
    If bSuccess And (g_Spells.c_PrioritySpellQueue.Count > 0) Then
        locDebug "OnSpellCastBegin: c_PrioritySpellQueue.Count > 0"
        Call g_Spells.c_PrioritySpellQueue.Pop ' Pop the spell off the queue
    End If
    
    'reset casting flag
    Call ResetCastingFlag(sSource, Not bSuccess)    'reset the ReadyToCast flag too if spell cast failed/got interupted

    'let the macro engine knows the current spell is done being casted
    If Valid(g_Macro) Then
        If g_Macro.Active Then
            Call g_Macro.OnSpellCasted(bSuccess)
        End If
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.OnSpellCastComplete - " & Err.Description
    Resume Fin
End Sub

Public Sub OnSpellCastFailed(Optional ByVal sSource As String = "")
    Call OnSpellCastComplete(False, sSource)
End Sub

' Priority Spell cast
Public Function CastPrioritySpell(ByRef spellList As clsSpellList, ByVal sSpellFamily As String, ByVal TargetId As Long, Optional ByVal iLevelWanted As Integer = 8) As Boolean
On Error GoTo ErrorHandler
    
    Dim objSpell As clsSpell
 
    'default return value
    CastPrioritySpell = False
    
    If Not m_bSpellsLoaded Then
        PrintErrorMessage "CastPrioritySpell: spells lists not loaded"
        Exit Function
    ElseIf Not Valid(spellList) Then
        PrintErrorMessage "CastPrioritySpell: invalid spelllist"
        Exit Function
    Else
        Set objSpell = spellList.FindSpell(sSpellFamily, iLevelWanted)
        If Valid(objSpell) Then
            CastPrioritySpell = CastThisSpell(objSpell, TargetId, True)
        Else
            PrintErrorMessage "(CastPrioritySpell) Couldn't cast " & sSpellFamily & " lvl " & iLevelWanted & " - Spell not found/learned."
            CastPrioritySpell = False
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.CastPrioritySpell - " & Err.Description
    Resume Fin
End Function

' Queues up a spell for later casting
Public Function CastThisSpell(ByVal objSpell As clsSpell, ByVal lTargetId As Long, Optional ByVal bPriority As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    
    CastThisSpell = False
        
    If Not Valid(objSpell) Then
        PrintErrorMessage "CastThisSpell: Invalid objSpell"
        Exit Function
    End If
    
    'remember when we tryed to cast our last spell
    m_dLastCastAttempt = g_Core.Time
    
    Call queueThisSpell(objSpell, lTargetId, bPriority)
    
    CastThisSpell = True
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.CastThisSpell - " & Err.Description
    Resume Fin
End Function

' Calls CastSpell if everything is OK
Public Function CastSpell(ByRef spellList As clsSpellList, ByVal sSpellFamily As String, ByVal TargetId As Long, Optional ByVal iLevelWanted As Integer = 8) As Boolean
On Error GoTo ErrorHandler
    
    Dim objSpell As clsSpell
 
    'default return value
    CastSpell = False
    
    If Not m_bSpellsLoaded Then
        PrintErrorMessage "CastSpell: spells lists not loaded"
        Exit Function
    ElseIf Not Valid(spellList) Then
        PrintErrorMessage "CastSpell: invalid spelllist"
        Exit Function
    Else
        Set objSpell = spellList.FindSpell(sSpellFamily, iLevelWanted)
        If Valid(objSpell) Then
            CastSpell = CastThisSpell(objSpell, TargetId)
        Else
            PrintErrorMessage "(CastSpell) Couldn't cast " & sSpellFamily & " lvl " & iLevelWanted & " - Spell not found/learned."
            CastSpell = False
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.CastSpell - " & Err.Description
    Resume Fin
End Function

' Finds the exact spell for a requested Type, Element, level, etc
Public Function CastSpellType(ByRef spellList As clsSpellList, ByVal iType As Integer, ByVal iElement As Integer, ByVal TargetId As Long, Optional ByVal iLevelWanted As Integer = 8) As Boolean
On Error GoTo ErrorHandler

    Dim objSpell As clsSpell
 
    'default return value
    CastSpellType = False
    
    If Not m_bSpellsLoaded Then
        PrintErrorMessage "CastSpellType: spells lists not loaded"
        Exit Function
    ElseIf Not Valid(spellList) Then
        PrintErrorMessage "CastSpellType: invalid spelllist"
        Exit Function
    Else
        Set objSpell = spellList.FindSpellByType(iType, iElement, iLevelWanted)
        If Valid(objSpell) Then
            CastSpellType = CastThisSpell(objSpell, TargetId)
        Else
            PrintErrorMessage "(CastSpellType) Couldn't cast " & GetSpelltypeString(iType) & " " & GetDamageString(iElement) & " " & iLevelWanted & " - Spell not found/learned."
            CastSpellType = False
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "ERROR: clsSpells.SpellCastType - " & Err.Description
    Resume Fin
End Function

Public Sub ResetCastingFlag(Optional sSource As String = "", Optional ByVal bResetReadyToCast As Boolean = False)
    locDebug "ResetCastingFlag - From: " & sSource
    m_bCasting = False
    'g_Macro.OkToCast = True
    Call m_tmrPrecastDelay.ExpireNow
    'Call g_Macro.PostCastDelayTimer.SetNextTime(0.5)
End Sub

Public Function CastVitalSpell(ByVal SpellBaseName As String, Optional ByVal SpellLevelWanted As Integer = 8, Optional ByVal TargetId As Long = 0) As Boolean
    'Call g_Spells.c_SpellQueue.Clear
    Call g_Spells.c_PrioritySpellQueue.Clear
    CastVitalSpell = CastPrioritySpell(m_Vitals, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function CastWarSpell(ByVal SpellBaseName As String, Optional ByVal SpellLevelWanted As Integer = 8, Optional ByVal TargetId As Long = 0) As Boolean
    CastWarSpell = CastSpell(m_War, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function CastWarType(ByVal iType As Integer, ByVal iElement As Integer, Optional ByVal SpellLevelWanted As Integer = 8) As Boolean
    CastWarType = CastSpellType(m_War, iType, iElement, SpellLevelWanted)
End Function

Public Function CastItemSpell(ByVal SpellBaseName As String, Optional SpellLevelWanted As Integer = 8, Optional ByVal TargetId As Long = 0) As Boolean
    CastItemSpell = CastSpell(m_Items, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function CastCreatureBuff(ByVal SpellBaseName As String, Optional ByVal SpellLevelWanted As Integer = 8, Optional ByVal TargetId As Long = 0) As Boolean
Dim spellList As clsSpellList
    
    If (TargetId <> 0) And (TargetId <> g_Objects.Player.Guid) Then
        Set spellList = m_oBuffs
    Else
        Set spellList = m_sBuffs
    End If
    
    CastCreatureBuff = CastSpell(spellList, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function CastLifeBuff(ByVal SpellBaseName As String, Optional SpellLevelWanted As Integer = 8, Optional ByVal TargetId As Long = 0) As Boolean
Dim spellList As clsSpellList
    
    If (TargetId <> 0) And (TargetId <> g_Objects.Player.Guid) Then
        Set spellList = m_oBuffs
    Else
        Set spellList = m_sBuffs
    End If
    
    CastLifeBuff = CastSpell(spellList, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function CastCreatureDebuff(ByVal SpellBaseName As String, Optional SpellLevelWanted As Integer = 7, Optional ByVal TargetId As Long = 0) As Boolean
    CastCreatureDebuff = CastSpell(m_Debuffs, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function CastLifeDebuff(ByVal SpellBaseName As String, Optional SpellLevelWanted As Integer = 7, Optional ByVal TargetId As Long = 0) As Boolean
    CastLifeDebuff = CastSpell(m_Debuffs, SpellBaseName, TargetId, SpellLevelWanted)
End Function

Public Function Cast_HealthToMana() As Boolean
    Cast_HealthToMana = CastVitalSpell("Health To Mana Self", g_Data.HealSpellsLevel)
End Function

Public Function Cast_Revitalize(Optional ByVal iLevel As Integer = -1) As Boolean
    If iLevel = -1 Then iLevel = g_Data.MacroSpellsLevel
    Cast_Revitalize = CastVitalSpell("Revitalize Self", iLevel)
End Function

Public Function Cast_StamToMana(Optional ByVal iLevel As Integer = -1) As Boolean
    If iLevel = -1 Then iLevel = g_Data.MacroSpellsLevel
    Cast_StamToMana = CastVitalSpell("Stamina To Mana Self", iLevel)
End Function

Public Function Cast_HealSelf() As Boolean
    Cast_HealSelf = CastVitalSpell("Heal Self", g_Data.HealSpellsLevel)
End Function

Public Function Cast_HealOther(ByVal aTargetID As Long) As Boolean
    Cast_HealOther = CastVitalSpell("Heal Other", g_Data.HealSpellsLevel, aTargetID)
End Function

Public Function Cast_StamOther(ByVal aTargetID As Long) As Boolean
    Cast_StamOther = CastVitalSpell("Revitalize Other", g_Data.HealSpellsLevel, aTargetID)
End Function


Public Function Cast_Emergency_Health2Mana() As Boolean
    'PrintMessage "Emergency Health To Mana Self " & g_Data.EmergencySpellsLevel
    Cast_Emergency_Health2Mana = CastVitalSpell("Health To Mana Self", g_Data.EmergencySpellsLevel)
End Function

Public Function Cast_Emergency_Stam2Health() As Boolean
    'PrintMessage "Emergency Stamina To Health Self " & g_Data.EmergencySpellsLevel
    Cast_Emergency_Stam2Health = CastVitalSpell("Stamina To Health Self", g_Data.EmergencySpellsLevel)
End Function

Public Function Cast_Emergency_Stam2Mana() As Boolean
    'PrintMessage "Emergency Stamina To Mana Self " & g_Data.EmergencySpellsLevel
    Cast_Emergency_Stam2Mana = Cast_StamToMana(g_Data.EmergencySpellsLevel)
End Function

Public Function Cast_Emergency_Revitalize() As Boolean
    PrintMessage "Emergency Revitalize Self " & g_Data.EmergencySpellsLevel
    Cast_Emergency_Revitalize = Cast_Revitalize(g_Data.EmergencySpellsLevel)
End Function

Public Function FindSpell(ByVal sSpellFamily As String, Optional ByVal iLevel As Integer = 8) As clsSpell
On Error GoTo ErrorHandler

    Dim objSpell As clsSpell

    Set objSpell = m_sBuffs.FindSpell(sSpellFamily, iLevel, , True)
    If Not Valid(objSpell) Then Set objSpell = m_Items.FindSpell(sSpellFamily, iLevel, , True)
    If Not Valid(objSpell) Then Set objSpell = m_oBuffs.FindSpell(sSpellFamily, iLevel, , True)
    If Not Valid(objSpell) Then Set objSpell = m_Debuffs.FindSpell(sSpellFamily, iLevel, , True)
    If Not Valid(objSpell) Then Set objSpell = m_War.FindSpell(sSpellFamily, iLevel, , True)
    
Fin:
    Set FindSpell = objSpell
    Set objSpell = Nothing
    Exit Function
ErrorHandler:
    Set objSpell = Nothing
    PrintErrorMessage "clsSpells.FindSpell - " & Err.Description
    Resume Fin
End Function

Public Function FindSpellByID(ByVal lSpellID As Long) As clsSpell
On Error GoTo ErrorHandler

    Dim objSpell As clsSpell

    Set objSpell = m_sBuffs.FindSpellByID(lSpellID)
    If Not Valid(objSpell) Then Set objSpell = m_Items.FindSpellByID(lSpellID)
    If Not Valid(objSpell) Then Set objSpell = m_oBuffs.FindSpellByID(lSpellID)
    If Not Valid(objSpell) Then Set objSpell = m_Debuffs.FindSpellByID(lSpellID)
    If Not Valid(objSpell) Then Set objSpell = m_War.FindSpellByID(lSpellID)
    
Fin:
    Set FindSpellByID = objSpell
    Set objSpell = Nothing
    Exit Function
ErrorHandler:
    Set objSpell = Nothing
    PrintErrorMessage "clsSpells.FindSpellByID - " & Err.Description
    Resume Fin
End Function

Public Function SetupWarSpellForDamage(ByVal iDmg As Integer, Optional fTargetSquareRange As Single = 0, Optional ByVal bSilent As Boolean = True, Optional ByVal overideDmg As Boolean = False) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim splArc As clsSpell
    Dim splWar As clsSpell
    Dim splRing As clsSpell
    
    'If we are using Auto switching rend wands, always match the damage type to the rend we are holding
    'As it should be correct (or the default damage)
    If overideDmg And g_ui.Macro.chkAutoSwitch.Checked And Valid(g_currentEquip) Then
        If g_currentEquip.DamageType <> DMG_NONE Then
            iDmg = g_currentEquip.DamageType
            locDebug "SetupWarSpellForDamage: overideDmg set iDmg to: " & iDmg
        End If
    End If
    
    'Force default will still overide the above check
    If g_ui.Macro.chkForceDefaultDmg.Checked Then
        iDmg = m_DefaultDamageType
        locDebug "SetupWarSpellForDamage: chkForceDefaultDmg set iDmg to: " & iDmg
    End If
    
    'default war spell type : bolt
    Set splWar = m_War.FindSpellByType(SPELLTYPE_BOLT, iDmg, g_Data.mMaxWarLevel)
    
    'Check if we can use an arc
    If g_ui.Macro.chkUseArcs.Checked Then
        Set splArc = m_War.FindSpellByType(SPELLTYPE_ARC, iDmg, g_Data.mMaxWarLevel)
        If Valid(splArc) Then
            If Valid(splWar) Then
                'See if target is at correct range to use an Arc
                'also compare spell levels between arc & bolt, keep the highest one
                If ((fTargetSquareRange >= (g_ui.Macro.txtArcRange.Text * g_ui.Macro.txtArcRange.Text)) And (splArc.SpellLevel >= splWar.SpellLevel)) _
                Or (splArc.SpellLevel > splWar.SpellLevel) Then
                    Set splWar = splArc
                End If
            Else
                Set splWar = splArc
            End If
        End If
    End If
    
    ' JSC: Ring spell support
    If g_ui.Macro.chkUseRings.Checked Then
        ' Need to check if we should use a ring spell here
        Set splRing = m_War.FindSpellByType(SPELLTYPE_RING, iDmg, g_Data.mMaxWarLevel)
        If Valid(splRing) Then
            ' Check to see if we are surrounded
            If FindRingSpellTarget(g_ui.Macro.txtRingRange.Text) Then
                Set splWar = splRing
                locDebug "SetupWarSpellForDamage: Using ring spell: " & splRing.SpellName
            End If
        Else
            PrintMessage "clsSpells: No valid RING spells for damage type: " & iDmg
        End If
    End If
     
    If Valid(splWar) Then
        Set m_CurrentWarSpell = splWar
        m_CurrentDamageType = m_CurrentWarSpell.SpellElement
        locDebug "SetupWarSpellForDamage: splWar: " & splWar.SpellName & " : " & splWar.SpellElement
        bRet = True
    Else
        Dim sMsg As String
        sMsg = "WARNING: Unable to find any " & GetDamageString(iDmg) & " bolt/arc spells. Using default."
        If bSilent Then
            MyDebug sMsg
        Else
            PrintMessage sMsg
        End If
        
        Set m_CurrentWarSpell = m_DefaultWarBolt
        m_CurrentDamageType = m_DefaultDamageType
        bRet = False
    End If
    
Fin:
    SetupWarSpellForDamage = bRet
    Set splArc = Nothing
    Set splWar = Nothing
    Set splRing = Nothing
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "SetupWarSpellForDamage - " & Err.Description
    Resume Fin
End Function

Public Function SetupVulnForDamage(ByVal iDmg As Integer, Optional ByVal bSilent As Boolean = True) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim splVuln As clsSpell
    
    If g_ui.Macro.chkForceDefaultDmg.Checked Then
        iDmg = m_DefaultDamageType
    End If
    
    Set splVuln = m_Debuffs.FindSpellByType(SPELLTYPE_VULN, iDmg, g_Data.mMaxVulnLevel)
    
    If Valid(splVuln) Then  'didnt even find default spell
        Set m_CurrentVuln = splVuln
        m_CurrentDamageType = m_CurrentVuln.SpellElement
        bRet = True
    Else
        Dim sMsg As String
        sMsg = "WARNING: Could not find a valid " & GetDamageString(iDmg) & " Vuln spell of level " & g_Data.mMaxVulnLevel & " (or under)."
        If bSilent Then
            MyDebug sMsg
        Else
            PrintMessage sMsg
        End If
        Set m_CurrentVuln = m_DefaultVuln
        m_CurrentDamageType = m_DefaultDamageType
        bRet = False
    End If

Fin:
    SetupVulnForDamage = bRet
    Set splVuln = Nothing
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "SetupVulnForDamage - " & Err.Description
    Resume Fin
End Function

'Returns false if couldn't find valid default/war vuln spells
Public Function SetDefaultDamage(ByVal iDmg As Integer, Optional sSource As String = "") As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim splWar As clsSpell
    Dim splVuln As clsSpell
    
    locDebug "SetDefaultDmg : " & iDmg & " (" & GetDamageString(iDmg) & ") - From: " & sSource
    
    'Default to false
    bRet = False
    
    'Try to find a valid life vuln
    Set splVuln = m_Debuffs.FindSpellByType(SPELLTYPE_VULN, iDmg, g_Data.mMaxVulnLevel)
    
    If Valid(splVuln) Then
        m_DefaultDamageType = iDmg
        Set m_DefaultVuln = splVuln
        
        If HasWarMagic Then
            Set splWar = m_War.FindSpellByType(SPELLTYPE_BOLT, iDmg, g_Data.mMaxWarLevel)
            If Not Valid(splWar) Then Set splWar = m_War.FindSpellByType(SPELLTYPE_ARC, iDmg, g_Data.mMaxWarLevel)
            
            If Valid(splWar) Then
                Set m_DefaultWarBolt = splWar
                bRet = True
            End If
        Else 'melee char without war - return ok
            bRet = True
        End If
    End If
    
Fin:
    SetDefaultDamage = bRet
    Set splWar = Nothing
    Set splVuln = Nothing
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsSpells.SetDefaultDamage - " & Err.Description
    Resume Fin
End Function

Public Function FindAndSetDefaultDamage() As Integer
On Error GoTo ErrorHandler

    Dim i As Integer
    Dim iRet As Integer

    'default if can't find
    iRet = DMG_NONE
    
    For i = DMG_SLASHING To DMG_LIGHTNING
        If SetDefaultDamage(i, "clsSpells.FindAndSetDefaultDamage") Then
            locDebug "clsSpells.FindAndSetDefaultDamage : found " & GetDamageString(i)
            iRet = i
            GoTo Fin
        End If
    Next i

Fin:
    FindAndSetDefaultDamage = iRet
    Exit Function
ErrorHandler:
    iRet = DMG_SLASHING
    PrintErrorMessage "clsSpells.FindAndSetDefaultDamage - " & Err.Description
    Resume Fin
End Function

'Public Function GetSpellName(ByVal lSpellID As Long) As String
'On Error GoTo ErrorHandler
'
'    Dim sRet As String
'    If m_dicSpellNames.Exists(lSpellID) Then
'        sRet = m_dicSpellNames(lSpellID)
'    Else
'        sRet = ""
'    End If
'
'Fin:
'    GetSpellName = sRet
'    Exit Function
'ErrorHandler:
'    sRet = ""
'    PrintErrorMessage "clsSpells.GetSpellName(" & lSpellID & ")"
'    Resume Fin
'End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   Private Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub queueThisSpell(ByVal objSpell As clsSpell, ByVal TargetGUID As Long, Optional ByVal bPriority As Boolean = False)
On Error GoTo ErrorHandler

1   If Not Valid(objSpell) Then
        PrintErrorMessage "queueThisSpell : Invalid objSpell"
        Exit Sub
    End If

    'Make sure we are in combat mode with a Wand
2   If Not g_Macro.ValidState(TYPE_CASTER) Then
        locDebug "queueThisSpell: Not in Combat mode - resetting"
    End If
    
'    'make sure we're able to cast
'3    If IsBusy(TurboMode) Then
'        MyDebug "queueThisSpell: trying to cast while busy - IsBusy()"
'        Exit Sub
'4    'ElseIf Not m_tmrPrecastDelay.Expired Then
'        MyDebug "queueThisSpell: tmrPrecastDelay not expired yet"
'        Exit Sub
'     Else
'        'PrintErrorMessage "queueThisSpell: casting " & SpellID & " at " & TargetGUID
'5       Call m_tmrPrecastDelay.SetNextTime(0.5)
'6       Call m_tmrSecureCasting.SetNextTime(SECURE_CASTING_TIME)
'     End if

    Dim s_objSpell As New clsSpellQueueItem
    With s_objSpell
        .SpellID = objSpell.SpellID
        .SpellFamily = objSpell.SpellFamily
        .SpellType = objSpell.SpellType
        .SpellElement = objSpell.SpellElement
        .SpellSchool = objSpell.SpellSchool
        .TargetGUID = TargetGUID
    End With
        
    If bPriority Then
        'First check to see if we already have this spell Queued up
        If c_PrioritySpellQueue.Count > 0 Then
            If c_PrioritySpellQueue.ReadTop.SpellID = s_objSpell.SpellID Then
                locDebug "queueThisSpell: we already have spell in PRIORITY queue: " & s_objSpell.Description
                Exit Sub
            End If
        End If
        MyDebug "clsSpells.queueThisSpell: adding PRIORITY objSpell: " & objSpell.SpellName & " :: " & objSpell.SpellID
        Call c_PrioritySpellQueue.Add(s_objSpell)
    Else
        'First check to see if we already have this spell Queued up
        If c_SpellQueue.Count > 0 Then
            If c_SpellQueue.ReadTop.SpellID = s_objSpell.SpellID Then
                locDebug "queueThisSpell: we already have spell in queue: " & s_objSpell.Description
                Exit Sub
            End If
        End If
        locDebug "clsSpells.queueThisSpell: adding objSpell: " & objSpell.SpellName & " :: " & objSpell.SpellID
        Call c_SpellQueue.Add(s_objSpell)
    End If
        
    'Call g_Hooks.CastSpell(objSpell.spellID, TargetGUID)
    'MyDebug "queueThisSpell: g_Hooks.CastSpell: " & objSpell.SpellID & " : " & TargetGUID
    
Fin:
    Set s_objSpell = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSpells.queueThisSpell " & Err.Description & "  line: " & Erl
    Resume Fin
End Sub

'Load the SpellNames.dat (data file with only spells ID/Name pairs
'Private Function LoadSpellNames(ByVal sPath As String) As Boolean
'On Error GoTo ErrorHandler
'    Dim bRet As Boolean
'
'    Dim db As New DataFile
'    Dim dat As clsDataEntry
'    Dim i As Integer
'
'    If db.Load(sPath) Then
'        For Each dat In db
'            i = i + 1
'            If dat.ParamExist(TAG_SPELL_ID) And dat.ParamExist(TAG_SPELL_NAME) Then
'                Dim lSpellID As Long
'                lSpellID = Val(dat.Param(TAG_SPELL_ID))
'                Call m_dicSpellNames.Add(lSpellID, dat.Param(TAG_SPELL_NAME))
'            End If
'        Next dat
'
'        MyDebug "LoadSpellNames : loaded " & m_dicSpellNames.Count & "/" & i & " spellnames"
'        bRet = True
'    Else
'        PrintErrorMessage "LoadSpellNames : failed to load SpellNames from " & sPath
'        bRet = False
'    End If
'
'Fin:
'    Set db = Nothing
'    Set dat = Nothing
'    LoadSpellNames = bRet
'    Exit Function
'ErrorHandler:
'    bRet = False
'    PrintErrorMessage "clsSpells.LoadSpellNames(" & sPath & ")"
'    Resume Fin
'End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Local Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub m_tmrCastBreaker_OnTimeout()
    If m_bCasting Then
        MyDebug "--- Cast Breaker ! ---"
        'Call g_Core.SendKeyHold(g_Keys.KeyJump)
        Call OnSpellCastComplete(, "tmrCastBreaker_OnTimeout")
        'Call g_Core.SendKeyRelease(g_Keys.KeyJump)
    End If
End Sub

Private Sub m_tmrSecureCasting_OnTimeout()
    If m_bCasting Then
        'Call g_Hooks.FaceHeading(g_Hooks.HeadingDegrees + 30, False)
        MyDebug "tmrSecureCasting timer expired. Resetting casting flags."
        Call ResetCastingFlag(, True)
    End If
End Sub

'
'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsSpells] " & DebugMsg, bSilent)
    End If
End Sub

