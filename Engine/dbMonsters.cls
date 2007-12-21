VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbMonsters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Monsters Info Database
Option Explicit

Private Const TAG_MOBNAME = "monster"
Private Const TAG_VULNERABILITY = "element"
Private Const TAG_VULN = "vuln"
Private Const TAG_YIELD = "yield"
Private Const TAG_IMPERIL = "imp"
Private Const TAG_PRIORITY = "priority"
Private Const TAG_ENABLED = "enabled"

Private m_db As DataFile
Private m_bLoaded As Boolean
Private m_colMonsters As Collection

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    MyDebug "dbMonster.Initialize - before New DataFile"
    Set m_db = New DataFile
    m_bLoaded = False
    
    MyDebug "dbMonster.Initialize - after New DataFile"
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "dbMonsters.Initialize - " & Err.Description & " - " & Err.Source
    Exit Sub
End Sub

Private Sub Class_Terminate()
    Set m_db = Nothing
    Set m_colMonsters = Nothing
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get Monsters() As Collection
    Set Monsters = m_colMonsters
End Property

'#####################################################################################
'#
'#                               PUBLIC
'#
'#####################################################################################

Public Function MonsterExist(ByVal bMonsterName As String) As Boolean
On Error GoTo NotFound
    Dim tmpMob As clsMonsterEntry
    Set tmpMob = m_colMonsters(bMonsterName)
    MonsterExist = True
Fin:
    Exit Function
NotFound:
    MonsterExist = False
    Resume Fin
End Function

Public Function AddMonster(ByVal bMonsterName As String) As clsMonsterEntry
On Error GoTo ErrorHandler

    Dim newMob As clsMonsterEntry
    
    If MonsterExist(bMonsterName) Then
        Set newMob = m_colMonsters(bMonsterName)
    Else
        Set newMob = New clsMonsterEntry
        newMob.MonsterName = bMonsterName
        Call m_colMonsters.Add(newMob, newMob.MonsterName)
    End If

    Set AddMonster = newMob
Fin:
    Exit Function
ErrorHandler:
    PrintWarning "dbMonsters.AddMonster(" & bMonsterName & ") - " & Err.Description
    Set AddMonster = Nothing
    Resume Fin
End Function

Public Function RemoveMonster(ByVal bMonsterName As String) As Boolean
    If MonsterExist(bMonsterName) Then
        Call m_colMonsters.Remove(bMonsterName)
        RemoveMonster = True
    Else
        RemoveMonster = False
    End If
End Function

Public Function FindMonster(ByVal bMonsterName As String, ByRef objMobOut As clsMonsterEntry) As Boolean
On Error GoTo NotFound
    
    Set objMobOut = m_colMonsters(bMonsterName)
    FindMonster = True

Fin:
    Exit Function
NotFound:
    Set objMobOut = Nothing
    FindMonster = False
    Resume Fin
End Function

Public Function LoadDatabase(sPath As String) As Boolean
On Error GoTo ErrorHandler

    m_bLoaded = m_db.Load(sPath)
    
    If Not m_bLoaded Then
        PrintErrorMessage "Failed to load Monsters Database from " & m_db.FileName
    Else
        Dim dat As clsDataEntry
        Dim mob As clsMonsterEntry
        
        Set m_colMonsters = New Collection
        
        For Each dat In m_db
            If Not dat.ParamExist(TAG_MOBNAME) _
            Or Not dat.ParamExist(TAG_VULNERABILITY) _
            Or Not dat.ParamExist(TAG_VULN) _
            Or Not dat.ParamExist(TAG_YIELD) _
            Or Not dat.ParamExist(TAG_IMPERIL) _
            Or Not dat.ParamExist(TAG_PRIORITY) _
            Or Not dat.ParamExist(TAG_ENABLED) Then
                PrintWarning "dbMonster.LoadDatabase : invalid monster entry. Missing one or more params."
            Else
                Set mob = AddMonster(dat.Param(TAG_MOBNAME))
                With mob
                    .MonsterVuln = GetDamageType(dat.Param(TAG_VULNERABILITY))
                    .Vuln = CBool(dat.Param(TAG_VULN))
                    .Yield = CBool(dat.Param(TAG_YIELD))
                    .Imperil = CBool(dat.Param(TAG_IMPERIL))
                    .Priority = CInt(dat.Param(TAG_PRIORITY))
                    .Enabled = CBool(dat.Param(TAG_ENABLED))
                End With
                
                'MyDebug "Loaded monster : " & mob.MonsterName, True
            End If
        Next dat
        
        Set dat = Nothing
        Set mob = Nothing
    End If
    
    MyDebug "[Database] Monsters loaded"
    
Fin:
    LoadDatabase = m_bLoaded
    Exit Function
ErrorHandler:
    PrintErrorMessage "dbMonsters.LoadDatabase - " & Err.Description
    m_bLoaded = False
    Resume Fin
End Function

Public Function SaveDatabase() As Boolean
On Error GoTo ErrorHandler

    If Not m_bLoaded Then
        PrintWarning "SaveDatabase : ignored, because database hasn't been loaded properly"
    Else
    
        Dim dat As clsDataEntry
        Dim mob As clsMonsterEntry
        
        'reset database content
        Call m_db.ResetData
        
        'Sort the DB before saving
        Call SortMonCollection(m_colMonsters)
        
        'write new database
        For Each mob In m_colMonsters
        
            Set dat = New clsDataEntry
            dat.AddParam TAG_MOBNAME, mob.MonsterName
            dat.AddParam TAG_VULNERABILITY, GetDamageString(mob.MonsterVuln)
            dat.AddParam TAG_VULN, BoolToInteger(mob.Vuln)
            dat.AddParam TAG_YIELD, BoolToInteger(mob.Yield)
            dat.AddParam TAG_IMPERIL, BoolToInteger(mob.Imperil)
            dat.AddParam TAG_PRIORITY, mob.Priority
            dat.AddParam TAG_ENABLED, BoolToInteger(mob.Enabled)
            
            If Not m_db.AddData(dat) Then
                PrintErrorMessage "....Failed to add " & mob.MonsterName
            End If
            
        Next mob
        
        If Not m_db.save Then
            PrintErrorMessage "Failed to save Monsters Database at " & m_db.FileName
            SaveDatabase = False
        Else
            MyDebug "Monster Database saved to " & m_db.FileName
            SaveDatabase = True
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "dbMonsters.SaveDatabase - " & Err.Description
    SaveDatabase = False
    Resume Fin
End Function

Public Sub SortMonCollection(ByRef ColVar As Collection)
On Error GoTo ErrorHandler
    
    Dim oCol As Collection
    Dim i As Integer
    Dim i2 As Integer
    Dim iBefore As Integer
    
    Dim mob As clsMonsterEntry
    Dim mob2 As clsMonsterEntry
    
    If ColVar.Count > 0 Then
        Set oCol = New Collection
        For i = 1 To ColVar.Count
            Set mob = ColVar(i)
            If oCol.Count = 0 Then
                Call oCol.Add(ColVar(i), mob.MonsterName)
            Else
                iBefore = 0
                For i2 = oCol.Count To 1 Step -1
                    Set mob2 = oCol(i2)
                    If LCase(mob.MonsterName) < LCase(mob2.MonsterName) Then
                        iBefore = i2
                    Else
                        Exit For
                    End If
                Next
                    
                If iBefore = 0 Then
                    Call oCol.Add(mob, mob.MonsterName)
                Else
                    Call oCol.Add(mob, mob.MonsterName, iBefore)
                End If
            End If
        Next
        
        Set ColVar = oCol
        Set oCol = Nothing
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SortMonCollection - " & Err.Description
    Resume Fin
End Sub

