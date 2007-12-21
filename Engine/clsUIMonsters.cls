VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIMonsters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================
'User Interface for tab : Monsters
'===========================================
Option Explicit

Private Const COL_MOB_ENABLED = 0
Private Const COL_MOB_NAME = 1
Private Const COL_MOB_VULNERABILITY = 3
Private Const COL_MOB_YIELD = 4
Private Const COL_MOB_IMPERIL = 5
Private Const COL_MOB_VULN = 6
Private Const COL_MOB_PRIORITY = 7

Private InterfaceName As String

Private m_curLineSelected As Long
Private m_PreviousSearchName As String
Private m_PreviousSearchPosition As Long
Private m_bListReady As Boolean

'Controls
Public WithEvents btnSearchMonster As PushButton
Attribute btnSearchMonster.VB_VarHelpID = -1
Public WithEvents btnDeleteMonster As PushButton
Attribute btnDeleteMonster.VB_VarHelpID = -1
Public WithEvents btnAddMonster As PushButton
Attribute btnAddMonster.VB_VarHelpID = -1
Public chMonsterVulnerability As DecalControls.Choice
Public WithEvents lstMonsters As DecalControls.list
Attribute lstMonsters.VB_VarHelpID = -1
Public WithEvents txtAddMonster As DecalControls.Edit
Attribute txtAddMonster.VB_VarHelpID = -1
Public chkAutoAddMonster As DecalControls.CheckBox
Public bListModified As Boolean

Private Sub Class_Terminate()
    Call Unload
End Sub

'Controls declaration
Public Function Init(Optional ProfileName As String = "Default") As Boolean

    On Error GoTo Error_Handler

    'Initialize
    Init = False
    
    'Set interface module name
    InterfaceName = "UIMonsters"
    
    MyDebug InterfaceName & ".Init() -- Begin"
    
    'Decal Controls initialisation
    Set lstMonsters = g_MainView.Control("lstMonsters")
    Set btnAddMonster = g_MainView.Control("btnAddMonster")
    Set chMonsterVulnerability = g_MainView.Control("chMonsterVulnerability")
    Set txtAddMonster = g_MainView.Control("txtAddMonster")
    Set btnSearchMonster = g_MainView.Control("btnSearchMonster")
    Set btnDeleteMonster = g_MainView.Control("btnDeleteMonster")
    Set chkAutoAddMonster = g_MainView.Control("chkAutoAddMonster")

    'Default Control Values
    m_curLineSelected = -1
    m_PreviousSearchName = ""
    m_PreviousSearchPosition = 0
    m_bListReady = False
    bListModified = False
    
    Init = True
    MyDebug InterfaceName & ".Init() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Init = False
    PrintErrorMessage InterfaceName & ".Init - " & Err.Description
    Resume Fin
End Function

Public Function LoadControlsValue(Optional ProfileName As String = "Default") As Boolean
On Error GoTo Error_Handler
    
    MyDebug "[" & InterfaceName & "] Loading controls value"
    
    m_curLineSelected = -1
    m_PreviousSearchName = ""
    m_PreviousSearchPosition = 0
    bListModified = False
    
    '***************************************************************************
    'Monster List Vulnerability
    chMonsterVulnerability.Selected = 0
    chkAutoAddMonster.Checked = g_Settings.GetValue("chkAutoAddMonster", True)
    '***************************************************************************
    
    'Fill the monsters list with the monster DB content
    Call LoadMonsterList
    
    LoadControlsValue = True
    
Fin:
    Exit Function
Error_Handler:
    LoadControlsValue = False
    PrintErrorMessage "(" & InterfaceName & " LoadControlsValue) - " & Err.Description
    Resume Fin

End Function

Public Function SaveControlSettings() As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    
    MyDebug "[" & InterfaceName & "] Saving controls settings"
    
    With g_Settings.Profile.Config
        .SaveCheckbox chkAutoAddMonster, "chkAutoAddMonster"
    End With
    
    bRet = True
Fin:
    SaveControlSettings = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage InterfaceName & ".SaveControlSettings - " & Err.Description
    Resume Fin
End Function

Public Function Unload() As Boolean
On Error GoTo Error_Handler
    
    MyDebug InterfaceName & ".Unload() -- Begin"
    
    'Unload initialization
    Unload = False

    Set lstMonsters = Nothing
    Set btnAddMonster = Nothing
    Set chMonsterVulnerability = Nothing
    Set txtAddMonster = Nothing
    Set btnSearchMonster = Nothing
    Set btnDeleteMonster = Nothing
    Set chkAutoAddMonster = Nothing
    
    'Unload complete
    Unload = True
    
    MyDebug InterfaceName & ".Unload() -- End"
    
Fin:
    Exit Function
Error_Handler:
    Unload = False
    PrintErrorMessage "(interface:" & InterfaceName & ") - " & Err.Description
    Resume Fin
End Function


Private Sub btnDeleteMonster_Accepted(ByVal nID As Long)

    If m_curLineSelected >= 0 Then
        Dim MonsterName As String
        MonsterName = lstMonsters.Data(COL_MOB_NAME, m_curLineSelected)
        
        If g_Data.mdbMonsters.RemoveMonster(MonsterName) Then
            PrintMessage MonsterName & " has been removed from the monsters database."
            lstMonsters.DeleteRow (m_curLineSelected)
            m_curLineSelected = -1
            bListModified = True
        Else
            PrintWarning "Unable to remove " & MonsterName & " from database. Reason: can't find monster entry in g_Data.mdbMonsters."
        End If
    Else
        PrintMessage "Please select a monster to remove from the list."
    End If
        
End Sub

Private Sub DoMonsterSearch(Optional ByVal bSilent As Boolean = True)
On Error GoTo ErrorHandler

    Dim SearchName As String
    Dim i As Integer, iStart As Integer

    If lstMonsters.Count = 0 Then
        Exit Sub
    End If
    
    SearchName = Trim(LCase(txtAddMonster.Text))
    
    If SearchName <> "" Then
        If m_PreviousSearchName = SearchName Then
            iStart = m_PreviousSearchPosition + 1
        Else
            iStart = 0
        End If

        'MyDebug "iStart=" & iStart
        If iStart < lstMonsters.Count Then
            For i = iStart To lstMonsters.Count - 1
                If InStr(1, LCase(lstMonsters.Data(COL_MOB_NAME, i)), SearchName) > 0 Then
                    'MyDebug "Found monster at line " & i
                    Call SelectMonsterLine(i, True)
                    m_PreviousSearchName = SearchName
                    m_PreviousSearchPosition = i
                    Exit Sub
                End If
            Next i
        End If
    Else
        If Not bSilent Then PrintMessage "Please enter the monster name to search in the text field."
        Exit Sub
    End If
    
    'No match found
    If m_PreviousSearchName = SearchName Then
        If Not bSilent Then PrintMessage "No other " & SearchName & " entries in the database."
        m_PreviousSearchPosition = -1 'move back to top
    Else
        If Not bSilent Then PrintMessage "Could not find " & txtAddMonster.Text & " in the monsters database."
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsUIMonsters.DoMonsterSearch - " & Err.Description
    Resume Fin
End Sub

Private Sub btnSearchMonster_Accepted(ByVal nID As Long)

    
    Call DoMonsterSearch(False)
    
    
End Sub


Private Sub lstMonsters_Change(ByVal nID As Long, ByVal nX As Long, ByVal nY As Long)
On Error GoTo Error_Handler

    Dim iVulnerability As Integer
    Dim sMobName As String
    Dim mobEntry As clsMonsterEntry

    If m_bListReady Then
    
        sMobName = lstMonsters.Data(COL_MOB_NAME, nY)
        
        If Not g_Data.mdbMonsters.FindMonster(sMobName, mobEntry) Then
            PrintWarning "lstMonsters_Change - Couldn't find monster entry for " & sMobName
            Exit Sub
        End If
        
        If Not Valid(mobEntry) Then
            PrintWarning "lstMonsters_Change - Invalid mobEntry"
            Exit Sub
        End If
        
        Select Case nX
            Case COL_MOB_ENABLED
                mobEntry.Enabled = Not mobEntry.Enabled
                bListModified = True
                
            Case COL_MOB_NAME
                Call SelectMonsterLine(nY)
                Exit Sub
            
            Case COL_MOB_VULNERABILITY
                'vulnerability column - cycle through the different damage type
                iVulnerability = mobEntry.MonsterVuln
                
                If iVulnerability >= DMG_LIGHTNING Then
                    iVulnerability = DMG_SLASHING
                Else
                    iVulnerability = iVulnerability + 1
                End If
                
                mobEntry.MonsterVuln = iVulnerability
                lstMonsters.Data(COL_MOB_VULNERABILITY, nY) = GetDamageString(iVulnerability)
                bListModified = True
                
            Case COL_MOB_YIELD
                mobEntry.Yield = Not mobEntry.Yield
                bListModified = True
                
            Case COL_MOB_IMPERIL
                mobEntry.Imperil = Not mobEntry.Imperil
                bListModified = True
                
            Case COL_MOB_VULN
                mobEntry.Vuln = Not mobEntry.Vuln
                bListModified = True
                            
            Case COL_MOB_PRIORITY
                Dim iPriority As Integer
                
                iPriority = mobEntry.Priority
                iPriority = iPriority + 1
                
                If iPriority >= 5 Then
                    iPriority = 1
                End If
                
                mobEntry.Priority = iPriority
                lstMonsters.Data(COL_MOB_PRIORITY, nY) = iPriority
                bListModified = True
        
            Case Else
                Exit Sub
                
        End Select
        
        'Update data for each mob in the world
        With mobEntry
            Call UpdateMonstersInWorld(.MonsterName, .MonsterVuln, .Vuln, .Yield, .Imperil, .Priority, .Enabled)
        End With
        
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage InterfaceName & "lstMonsters_Change - " & Err.Description
    Resume Fin
End Sub

Public Sub SelectMonsterLine(ByVal LineIndex As Integer, Optional MoveToSelection As Boolean = False)
On Error GoTo ErrorHandler

    If LineIndex < 0 Or LineIndex >= lstMonsters.Count Then
        Exit Sub
    End If
    
    'unselect the previous line
    If m_curLineSelected >= 0 Then
        lstMonsters.Color(COL_MOB_NAME, m_curLineSelected) = vbWhite
        lstMonsters.Color(COL_MOB_VULNERABILITY, m_curLineSelected) = vbWhite
        lstMonsters.Color(COL_MOB_PRIORITY, m_curLineSelected) = vbWhite
    End If
    
    'select new line
    m_curLineSelected = LineIndex
    lstMonsters.Color(COL_MOB_NAME, m_curLineSelected) = vbGreen
    lstMonsters.Color(COL_MOB_VULNERABILITY, m_curLineSelected) = vbGreen
    lstMonsters.Color(COL_MOB_PRIORITY, m_curLineSelected) = vbGreen
    
    If MoveToSelection Then lstMonsters.JumpToPosition (m_curLineSelected)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage InterfaceName & ".btnAddMonster_Accepted - " & Err.Description
    Resume Fin
    
End Sub

Private Sub btnAddMonster_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    Dim i As Long
    Dim mobEntry As clsMonsterEntry

    Set mobEntry = g_Data.mdbMonsters.AddMonster(txtAddMonster.Text)
        
    If Not Valid(mobEntry) Then
        PrintErrorMessage "btnAddMonster_Accepted : invalid mobEntry"
        Exit Sub
    End If
    
    With mobEntry
        .MonsterVuln = chMonsterVulnerability.Selected
        .Vuln = True
        .Imperil = True
        .Yield = True
        .Priority = 1
        .Enabled = True
        i = AddMonster(.MonsterName, .MonsterVuln, .Yield, .Imperil, .Vuln, .Priority, .Enabled)
    End With
    
    Call SelectMonsterLine(i, True)
    bListModified = True

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage InterfaceName & ".btnAddMonster_Accepted - " & Err.Description
    Resume Fin
    
End Sub

Public Function AddMonster(ByVal MonsterName As String, _
        ByVal MonsterVulnerability As Integer, _
        ByVal MonsterYield As Boolean, _
        ByVal MonsterImperil As Boolean, _
        ByVal MonsterVuln As Boolean, _
        ByVal MonsterPriority As Integer, _
        Optional ByVal MonsterEnabled As Boolean = True) As Long
        
On Error GoTo Error_Handler

    Dim curMonsterName As String
    Dim i As Integer
    
    AddMonster = -1
    
    If MonsterName = "" Then
        Exit Function
    End If
    
    For i = 0 To lstMonsters.Count - 1
        If SameText(lstMonsters.Data(COL_MOB_NAME, i), MonsterName) Then
            PrintMessage MonsterName & " already in the list."
            Exit Function
        End If
    Next i


'ATTENTION !
'Macro.Monster(i) = List(index = i)
'ex: listControl.data(COL_MOB_NAME, 0) = Macro.Monster(0)

    'find insertion position
'    i = -1
'    If listControl.Count > 0 Then
'        For j = 0 To listControl.Count - 1
'            curMonsterName = listControl.Data(COL_MOB_NAME, j)
'
'            If SameText(curMonsterName, MonsterName) Then
'                PrintMessage MonsterName & " already in the list."
'                Exit Function
'            End If
'
'            If StrComp(curMonsterName, MonsterName, vbTextCompare) > 0 Then
'                i = j
'                Exit For
'            End If
'        Next j
'    End If
'
'    If i = -1 Then
'        i = listControl.AddRow
'    Else
'        Call listControl.InsertRow(i)
'    End If
    

    i = lstMonsters.AddRow
    lstMonsters.Data(COL_MOB_ENABLED, i) = MonsterEnabled
    lstMonsters.Data(COL_MOB_NAME, i) = MonsterName
    lstMonsters.Data(COL_MOB_VULNERABILITY, i) = GetDamageString(MonsterVulnerability)
    lstMonsters.Data(COL_MOB_YIELD, i) = MonsterYield
    lstMonsters.Data(COL_MOB_IMPERIL, i) = MonsterImperil
    lstMonsters.Data(COL_MOB_VULN, i) = MonsterVuln
    lstMonsters.Data(COL_MOB_PRIORITY, i) = MonsterPriority
    
    AddMonster = i
    
Fin:
    Exit Function
Error_Handler:
    PrintErrorMessage InterfaceName & ".AddMonster - " & Err.Description
    Resume Fin
End Function

Private Sub LoadMonsterList()
On Error GoTo ErrorHandler

    Dim mobEntry As clsMonsterEntry
    Dim i As Long

    MyDebug InterfaceName & ".LoadMonsterList - Loading MonsterList"
    
    m_bListReady = False
        
    If Not Valid(g_Data) Then
        PrintWarning "LoadMonsterList : invalid g_Data"
        Exit Sub
    End If
    
    Call lstMonsters.Clear
    For Each mobEntry In g_Data.mdbMonsters.Monsters
        With mobEntry
            Call AddMonster(.MonsterName, .MonsterVuln, .Yield, .Imperil, .Vuln, .Priority, .Enabled)
        End With
    Next mobEntry

    m_bListReady = True
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage InterfaceName & ".LoadMonsterList - " & Err.Description
    Resume Fin
    
End Sub

Public Sub SaveMonsters()
    If bListModified And Valid(g_Data) Then
        MyDebug "[" & InterfaceName & "] Saving Monster List"
        Call g_Data.mdbMonsters.SaveDatabase
    End If
End Sub

Private Sub txtAddMonster_Change(ByVal nID As Long, ByVal strText As String)
    If strText <> "" Then
        Call DoMonsterSearch
    End If
End Sub

