Attribute VB_Name = "Monsters"
Option Explicit

'------------------------------------------------------------------
' InitMonster
'------------------------------------------------------------------
'
'   Called after a monster object is created (CreateObject filter
'   event in clsACEvents)
'
'   Initializes the extra monsters info required by the macro,
'   such as monster vulnerability, and adds new monsters to the
'   database when found
'
'------------------------------------------------------------------
Public Sub InitMonster(ByVal objMonster As acObject)
On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim oMobData As clsMonsterEntry
    
    If Not Valid(objMonster) Then
        PrintErrorMessage "InitMonsters - Invalid objMonster"
        Exit Sub
    End If
    
    'Set Default Settings
    With objMonster
        .Priority = 0 'Lowest priority
        .Vulnerability = g_Spells.DefaultDamage
        
        Call .SetUserData(B_ENABLED, True)
        Call .SetUserData(B_CAN_BE_IMPERILED, True)
        Call .SetUserData(B_CAN_BE_VULNED, True)
        Call .SetUserData(B_CAN_BE_YIELDED, True)
        Call .SetUserData(INT_YIELD_TRYS_LEFT, 3)
        Call .SetUserData(B_DANGEROUS, False)
        Call .SetUserData(B_WAR_CHECK, False)
        Call .SetUserData(INT_MISSCOUNT, 0)
        Call .SetUserData(INT_BLISTCOUNT, 0)
        Call .SetUserData(INT_BLIST_TIME, g_Core.Time)
    
        'Try to find more detailed info from the Monsters Database and complete where needed
        If g_Data.mdbMonsters.FindMonster(objMonster.Name, oMobData) Then
        
            .Priority = oMobData.Priority 'Lowest priority
            .Vulnerability = oMobData.MonsterVuln
            Call .SetUserData(B_ENABLED, oMobData.Enabled)
            Call .SetUserData(B_CAN_BE_IMPERILED, oMobData.Imperil)
            Call .SetUserData(B_CAN_BE_VULNED, oMobData.Vuln)
            Call .SetUserData(B_CAN_BE_YIELDED, oMobData.Yield)
        
        'This is a new monster, not in database - Auto-Add monster if option checked
        ElseIf g_ui.Monsters.chkAutoAddMonster.Checked Then
        
            PrintMessage "New monster detected : " & objMonster.Name & " - Auto-adding to Monsters Database."
            Set oMobData = g_Data.mdbMonsters.AddMonster(objMonster.Name)
            
            oMobData.MonsterVuln = .Vulnerability
            oMobData.Priority = .Priority
            oMobData.Vuln = CBool(.UserData(B_CAN_BE_VULNED))
            oMobData.Yield = CBool(.UserData(B_CAN_BE_YIELDED))
            oMobData.Imperil = CBool(.UserData(B_CAN_BE_IMPERILED))
            oMobData.Enabled = CBool(.UserData(B_ENABLED))
            
            Call g_ui.Monsters.AddMonster(objMonster.Name, _
                                            oMobData.Vuln, _
                                            oMobData.Yield, _
                                            oMobData.Imperil, _
                                            oMobData.Vuln, _
                                            oMobData.Priority, _
                                            oMobData.Enabled)
            
            'Set monster list as modified so that it will be saved at logout
            g_ui.Monsters.bListModified = True
        End If
    End With

Fin:
    Set oMobData = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "InitMonster - " & Err.Description
    Resume Fin
End Sub

'------------------------------------------------------------------
' UpdateMonstersInWorld
'------------------------------------------------------------------
'
'   Called after the user modified the monsters database settings
'   in the plugin's Monsters tab (clsUIMonsters)
'
'   It goes through the monsters currently alive in the world and
'   update their info to match the database changes on the fly
'
'   Note:   Only updates info of the mobs with name matching
'           sFullMonsterName
'
'------------------------------------------------------------------
Public Sub UpdateMonstersInWorld(ByVal sFullMonsterName As String, ByVal iVulnerability As Integer, _
                                ByVal bCanBeVulned As Boolean, ByVal bCanBeYielded As Boolean, _
                                ByVal bCanBeImperiled As Boolean, _
                                ByVal iPriority As Integer, ByVal bEnabled As Boolean)
On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim objMonster As acObject
    
    sFullMonsterName = Trim(LCase(sFullMonsterName))
    
    For Each objMonster In g_Objects.Monsters
        If SameText(sFullMonsterName, objMonster.Name) Then
            
            MyDebug "UpdateMonstersInWorld : updating monster " & objMonster.Name
            With objMonster
                .Priority = iPriority 'Lowest priority
                .Vulnerability = iVulnerability
                Call .SetUserData(B_ENABLED, bEnabled)
                Call .SetUserData(B_CAN_BE_IMPERILED, bCanBeImperiled)
                Call .SetUserData(B_CAN_BE_VULNED, bCanBeVulned)
                Call .SetUserData(B_CAN_BE_YIELDED, bCanBeYielded)
            End With
            
        End If
    Next objMonster
    
Fin:
    Set objMonster = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "UpdateMonstersInWorld - " & Err.Description
    Resume Fin
End Sub

