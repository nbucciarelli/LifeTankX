VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------
' Stores all the Damage Over Time variables and info
'-------------------------------------------------------

Private m_takeDamage As Collection
Private m_takeSpellDamage As Collection

Private m_giveSpellDamage As Collection
Private m_giveMeleeDamage As Collection

Private m_v3dText As Single
Private m_dMark As CD3DObj
Private WithEvents m_tmr3dText2 As Timer
Attribute m_tmr3dText2.VB_VarHelpID = -1
Private WithEvents m_tmr3dText As clsTimer
Attribute m_tmr3dText.VB_VarHelpID = -1

'----------------------------------------------------------------------------
' Init routines
'----------------------------------------------------------------------------

Private Sub Class_Initialize()
    Set m_takeDamage = New Collection
    Set m_takeSpellDamage = New Collection
    
    Set m_giveSpellDamage = New Collection
    Set m_giveMeleeDamage = New Collection
    
    'Set m_tmr3dText2 = frmTimer.tmr3Dtext
    'm_tmr3dText2.Enabled = False
    Set m_tmr3dText = CreateTimer
End Sub

Private Sub Class_Terminate()
    Set m_takeSpellDamage = Nothing
    Set m_takeDamage = Nothing

    Set m_giveSpellDamage = Nothing
    Set m_giveMeleeDamage = Nothing
    
    'Set m_tmr3dText2 = Nothing
    Set m_tmr3dText = Nothing
    Set m_dMark = Nothing
End Sub

Public Sub Init()
    MyDebug "clsDOT.Init called"
    Set m_dMark = g_D3D.m_d3ds.MarkObjectWith3DText(g_ds.Player.Guid, 0, "Arial", 0)
    Call m_dMark.SetScale(0.3)
    m_dMark.autoscale = False
    m_dMark.Color = &HCCBB0055
    m_dMark.Visible = False
    m_v3dText = 0
End Sub

Public Sub Reset()
    Set m_takeSpellDamage = New Collection
    Set m_takeDamage = New Collection

    Set m_giveSpellDamage = New Collection
    Set m_giveMeleeDamage = New Collection
End Sub


'------------------------------------
' Public Functions
'------------------------------------

Public Sub takeSpellDamage(ByVal aMsg As String)
On Error GoTo ErrorHandler

    Dim dObj As clsDOTobj
    Dim aLoc As String
    Dim aType As String
    Dim aName As String
    Dim aVal As Long
    Dim isCrit As Boolean
    Dim aRegex As New RegExp
    Dim colMatches As MatchCollection
    Dim m As match
    Dim aMatch As String
    
    Dim nSplit() As String
    Dim i, iArgs, iCount As Integer
    
    aMsg = LCase(aMsg)
    
    aRegex.Global = True
    aRegex.IgnoreCase = True
    
    Set dObj = Nothing
    
    MyDebug "cslDOT.takeSpellDamage: " & aMsg
    
    If (InStr(aMsg, "you for") And InStr(aMsg, "points with")) Then
    
        'White Phyntos Wasp chills you for 21 points with Frost Bolt IV.
        'Ripper Grievver shocks you for 63 points with Alset's Coil.
        'Ripper Grievver sears you for 51 points with Disintegration.

            'You resist the spell cast by White Phyntos Wasp
            
        ' {Name} chills you for {aVal} points with {Spell Name}

        If InStr(aMsg, ".") Then
            aMsg = Replace(aMsg, ".", "")
            isCrit = True
        End If

        ' remove Critcal message
        If InStr(aMsg, "critical") Then
            aMsg = Replace(aMsg, "critical hit! ", "")
            isCrit = True
        End If
        
        ' Get name
        nSplit = Split(aMsg, " ")
        iArgs = UBound(nSplit)
        For i = 0 To iArgs
            If InStr(nSplit(i), "you") Then
                iCount = (i - 2)
                Exit For
            End If
        Next i
        
        For i = 0 To iCount
            aMatch = aMatch & " " & nSplit(i)
        Next i
        
        aName = Trim(aMatch)
        
        ' get damage number
        aRegex.Pattern = " for (\d+) points "
        Set colMatches = aRegex.Execute(aMsg)
        aMatch = colMatches.Item(0).SubMatches(0)
        aVal = CLng(aMatch)
    
        ' get damage type
        'aRegex.Pattern = " points with (\w+)"
        'Set colMatches = aRegex.Execute(aMsg)
        'aMatch = colMatches.Item(0).SubMatches(0)
        iArgs = InStr(aMsg, "points with ")
        aMatch = Right(aMsg, (Len(aMsg) - (iArgs + 11)))
        aType = Trim(aMatch)
    
        MyDebug "clsDOT.takeSpellDamage: aName: " & aName & " dmg: " & aVal & " aType: " & aType
        
        ' Get damage location
        'aRegex.Pattern = " your (\w+) for "
        'Set colMatches = aRegex.Execute(aMsg)
        'aMatch = colMatches.Item(0).SubMatches(0)
        'aLoc = Trim(aMatch)
    
        If Exists(m_takeSpellDamage, aName) Then
            Set dObj = m_takeSpellDamage.Item(aName)
        Else
            Set dObj = New clsDOTobj
            Call dObj.setName(aName)
            Call m_takeSpellDamage.Add(dObj, aName)
        End If
    
        If Valid(dObj) Then
            Call dObj.damageByType(aType, aVal)
        End If

    ElseIf InStr(aMsg, "you resist the spell") Then

        'You resist the spell cast by White Phyntos Wasp

        ' get name
        'aMatch = Replace(aMsg, "!", "")
        aMatch = Replace(aMatch, "you resist the spell cast by", "")
        aName = Trim(aMatch)
        
        If InStr(aName, vbLf) Then
            aName = Replace(aName, vbLf, "")
            MyDebug "Detected vbLf in aName: " & aName
        End If
    
        MyDebug "clsDOT.meResisted: aName: " & aName

        If Exists(m_takeSpellDamage, aName) Then
            Set dObj = m_takeSpellDamage.Item(aName)
            Call dObj.addExtra
        Else
            Set dObj = New clsDOTobj
            Call dObj.setName(aName)
            Call dObj.addExtra
            Call m_takeSpellDamage.Add(dObj, aName)
        End If
        
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.takeSpellDamage - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub


Public Sub takeMeleeDamage(ByVal sAttacker As String, ByVal lDamage As Long, ByVal bCrit As Boolean, ByVal sDamageType As String, ByVal sLocation As String)
On Error GoTo ErrorHandler
    
    Dim dObj As clsDOTobj
    
    sAttacker = LCase(sAttacker)
        
    'MyDebug "clsDOT.takeMeleeDamage: aName: " & sAttacker & " dmg: " & lDamage & " aType: " & sDamageType
        
    If Exists(m_takeDamage, sAttacker) Then
        Set dObj = m_takeDamage.Item(sAttacker)
    Else
        Set dObj = New clsDOTobj
        Call dObj.setName(sAttacker)
        Call m_takeDamage.Add(dObj, sAttacker)
    End If
    
    If Valid(dObj) Then
        Call dObj.damageByType(sDamageType, lDamage)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.takeMeleeDamage - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

Public Sub playerEvade(ByVal sAttacker As String)
On Error GoTo ErrorHandler
    
    Dim dObj As clsDOTobj
    
    sAttacker = LCase(sAttacker)
        
    'MyDebug "clsDOT.meEvaded: aName: " & sAttacker

    If Exists(m_takeDamage, sAttacker) Then
        Set dObj = m_takeDamage.Item(sAttacker)
        Call dObj.addExtra
    Else
        Set dObj = New clsDOTobj
        Call dObj.setName(sAttacker)
        Call dObj.addExtra
        Call m_takeDamage.Add(dObj, sAttacker)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.playerEvade - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub


'----------------------------------------------------

Public Sub giveMeleeDamage(ByVal sTarget As String, ByVal lDamage As Long, ByVal bCrit As Boolean, ByVal sDamageType As String)
On Error GoTo ErrorHandler

    Dim dObj As clsDOTobj
    Dim aWeap As String
   
    sTarget = LCase(sTarget)
    
    If IsMelee Then
        ' Get name of weapon
        If g_Macro.CombatType = TYPE_MELEE Then
            aWeap = LCase(g_Data.Weapon.Name)
        Else
            aWeap = LCase(g_Data.Bow.Name)
        End If
    Else
        aWeap = "nothing"
    End If

    'MyDebug "clsDOT.giveMeleeDamage: aWeap: " & aWeap & " dmg: " & lDamage

        If g_ui.Options.chk3DTarget.Checked And Valid(g_Macro.Combat.Target) Then
            'Put some text above the target
            Call m_dMark.AnchorToObject(g_Macro.Combat.Target.Guid, 1.1, 0, 0, 0)
            Call m_dMark.Set3DText(lDamage, "Arial", 0)
            m_dMark.Visible = True
            m_v3dText = 0
            Call m_tmr3dText.SetNextTime(3)
            m_tmr3dText.Enabled = True
        End If

    If Exists(m_giveMeleeDamage, aWeap) Then
        Set dObj = m_giveMeleeDamage.Item(aWeap)
    Else
        Set dObj = New clsDOTobj
        Call dObj.setName(aWeap)
        Call m_giveMeleeDamage.Add(dObj, aWeap)
    End If
    
    If Valid(dObj) Then
        Call dObj.damageByType(sTarget, lDamage)
        If bCrit Then Call dObj.addExtra
    End If
        
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.giveMeleeDamage - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

Public Sub targetEvade(ByVal sTarget As String)
On Error GoTo ErrorHandler
    
    Dim dObj As clsDOTobj
    Dim aWeap As String
    
    sTarget = LCase(sTarget)
    
    If IsMelee Then
        ' Get name of weapon
        If g_Macro.CombatType = TYPE_MELEE Then
            aWeap = LCase(g_Data.Weapon.Name)
        Else
            aWeap = LCase(g_Data.Bow.Name)
        End If
    Else
        aWeap = "nothing"
    End If
        
    'MyDebug "clsDOT.targetEvade: aName: " & sTarget

    If Exists(m_giveMeleeDamage, aWeap) Then
        Set dObj = m_giveMeleeDamage.Item(aWeap)
    Else
        Set dObj = New clsDOTobj
        Call dObj.setName(aWeap)
        Call m_giveMeleeDamage.Add(dObj, aWeap)
    End If
    
    If Valid(dObj) Then
        Call dObj.extraByName(sTarget)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.targetEvade - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub


Public Sub giveSpellDamage(ByVal aMsg As String)
On Error GoTo ErrorHandler

    Dim dObj As clsDOTobj
    Dim SpellName As String
    Dim aWand As String
    Dim aName As String
    Dim aVal As Long
    Dim isCrit As Boolean
    Dim aRegex As New RegExp
    Dim colMatches As MatchCollection
    Dim m As match
    Dim aMatch As String
   
    Dim nSplit() As String
    Dim i, iArgs, iCount As Integer
    
    aMsg = LCase(aMsg)
    
    aRegex.Global = True
    aRegex.IgnoreCase = True
    
    ' Get name of wand
    If g_Data.Wand.Equiped Then
            aWand = LCase(g_Data.Wand.Name)
    Else
            aWand = "nothing"
    End If
    
    If InStr(aMsg, "points with") Then
    
        'Critical hit! You bash Olthoi Ripper for 786 points with Shock Arc VII.
        'You bash Olthoi Ripper for 516 points with Shock Arc VII.
    
        'You bash Olthoi Ripper for {aVal} points with {spellName}.
    
        ' remove Critcal message
        If InStr(aMsg, "critical") Then
            aMsg = Replace(aMsg, "critical hit! ", "")
            isCrit = True
        End If
        
        ' Get name
        nSplit = Split(aMsg, " ")
        iArgs = UBound(nSplit)
        For i = 0 To iArgs
            If InStr(nSplit(i), "for") Then
                iCount = (i - 1)
                Exit For
            End If
        Next i
        
        For i = 2 To iCount
            aMatch = aMatch & " " & nSplit(i)
        Next i
        
        ' get target name
        'aRegex.Pattern = "you (?:\w) (\w+) for "
        'Set colMatches = aRegex.Execute(aMsg)
        'For Each m In colMatches
        '    MyDebug "match: " & m.Value
        '   aMatch = aMatch & " " & m.Value
        'Next
        'aMatch = colMatches.Item(0).SubMatches(0)
        aName = Trim(aMatch)
    
        ' get damage number
        aRegex.Pattern = " for (\d+) points "
        Set colMatches = aRegex.Execute(aMsg)
        aMatch = colMatches.Item(0).SubMatches(0)
        aVal = CLng(aMatch)
   
        'MyDebug "clsDOT.giveSpellDamage: aWand: " & aWand & " aName: " & aName & " dmg: " & aVal

        If g_ui.Options.chk3DTarget.Checked And Valid(g_Macro.Combat.Target) Then
            'Put some text above the target
            Call m_dMark.AnchorToObject(g_Macro.Combat.Target.Guid, 1.1, 0, 0, 0)
            Call m_dMark.Set3DText(aVal, "Arial", 0)
            m_dMark.Visible = True
            m_v3dText = 0
            Call m_tmr3dText.SetNextTime(3)
            m_tmr3dText.Enabled = True
        End If

        If Exists(m_giveSpellDamage, aWand) Then
            Set dObj = m_giveSpellDamage.Item(aWand)
        Else
            Set dObj = New clsDOTobj
            Call dObj.setName(aWand)
            Call m_giveSpellDamage.Add(dObj, aWand)
        End If
    
        If Valid(dObj) Then
            Call dObj.damageByType(aName, aVal)
            If isCrit Then Call dObj.addExtra
        End If
    
    ElseIf InStr(aMsg, "resists your spell") Then
    
        'Olthoi Ripper resists your spell
        
        aMatch = Replace(aMsg, "!", "")
        aMatch = Replace(aMatch, "resists your spell", "")
        aName = Trim(aMatch)
        
        ' get name of critter
        'aRegex.Pattern = "(\w+) resists your spell"
        'Set colMatches = aRegex.Execute(aMsg)
        'aMatch = colMatches.Item(0).SubMatches(0)
        'aName = Trim(aMatch)
        
        'MyDebug "clsDOT.giveSpellDamage Resist: aWand: " & aWand & " aName: " & aName
        
        If g_ui.Options.chk3DTarget.Checked And Valid(g_Macro.Combat.Target) Then
            'Put some text above the target
            Call m_dMark.AnchorToObject(g_Macro.Combat.Target.Guid, 1.1, 0, 0, 0)
            Call m_dMark.Set3DText("resist", "Arial", 0)
            m_dMark.Visible = True
            m_v3dText = 0
            Call m_tmr3dText.SetNextTime(3)
            m_tmr3dText.Enabled = True
        End If
        
        If Exists(m_giveSpellDamage, aWand) Then
            Set dObj = m_giveSpellDamage.Item(aWand)
        Else
            Set dObj = New clsDOTobj
            Call dObj.setName(aWand)
            Call m_giveSpellDamage.Add(dObj, aWand)
        End If
    
        If Valid(dObj) Then
            Call dObj.extraByName(aName)
        End If
    
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.giveSpellDamage - " & Err.Description & " - line:" & Erl
    Resume Fin
End Sub

'----------------------------------------------------
' 3D Text Floating timer with it's own clock source
' gets called every 1/10 of a second
Private Sub m_tmr3dText2_Timer()
On Error GoTo ErrorHandler
    
    If m_v3dText < 0.3 And Valid(g_Macro.Combat.Target) Then
        m_v3dText = m_v3dText + 0.01
        Call m_dMark.AnchorToObject(g_Macro.Combat.Target.Guid, 1.3, 0, 0, m_v3dText)
    Else
        m_dMark.Visible = False
        m_v3dText = 0
        m_tmr3dText2.Enabled = False
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsDOT.m_tmr3dText2_Timer - " & Err.Description
    Resume Fin
End Sub

'Make combat numbers invisible
Private Sub m_tmr3dText_OnTimeout()
    m_dMark.Visible = False
End Sub


'----------------------------------------------------

Public Function colTakeMeleeDamage() As Collection
    Set colTakeMeleeDamage = m_takeDamage
End Function

Public Function colTakeSpellDamage() As Collection
    Set colTakeSpellDamage = m_takeSpellDamage
End Function

Public Function colGiveMeleeDamage() As Collection
    Set colGiveMeleeDamage = m_giveMeleeDamage
End Function

Public Function colGiveSpellDamage() As Collection
    Set colGiveSpellDamage = m_giveSpellDamage
End Function

'----------------------------------------------------
' Private Functions
'----------------------------------------------------

'Default Class Method
Private Function Exists(ByVal aCollection As Collection, ByVal aKey As String) As Boolean
On Error GoTo NotFound
    Dim anObj As clsDOTobj
    Set anObj = aCollection.Item(aKey)
    Exists = True
Fin:
    Exit Function
NotFound:
    Exists = False
    Resume Fin
End Function
