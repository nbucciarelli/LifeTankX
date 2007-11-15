Attribute VB_Name = "Vitals"
Option Explicit

'Returns true if we need to regen stamina - if bRestamMelee is set, it means
'we're in melee mode and checking to see if our stamina dropped too low
Public Function NeedStamina(Optional bRestamMelee As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    
    Dim iPercent As Integer
    
    If bRestamMelee Then
        iPercent = RESTAM_MIN_STAM_PERCENT
    Else
        iPercent = g_Data.MinStamThreshold
    End If
    
    NeedStamina = (g_Filters.Stam <= GetPercent(g_Filters.MaxStam, iPercent))
    
Fin:
    Exit Function
ErrorHandler:
    NeedStamina = False
    PrintErrorMessage "Vitals.NeedStamina - " & Err.Description
    Resume Fin
End Function


'***************************************
' CastManaStamRegen
'
' In magic mode, check if the macro needs to regen
' its mana/stamina. Returns true if it does, and
' casts the appropriate spell (revit/stam2mana, etc)
'***************************************
Public Function CastManaStamRegen(Optional bMeleeRestam As Boolean = False, _
                                    Optional ByVal bManaRegenMethod As eRegenActions = REGEN_REVITALIZE, _
                                    Optional ByVal bStamRegenMethod As eRegenActions = REGEN_REVITALIZE) As Boolean
On Error GoTo ErrorHandler

    Dim StamThresholdPercent As Integer
    Dim bRes As Boolean

    Dim curHealth As Integer
    Dim curMaxHealth As Integer
    Dim curMana As Integer
    Dim curMaxMana As Integer
    Dim curStam As Integer
    Dim curMaxStam As Integer
    
    If Not g_Macro.ValidState(TYPE_CASTER) Then
        bRes = False
        GoTo Fin
    End If
    
    With g_Filters
        curHealth = .Health
        curMana = .Mana
        curStam = .Stam

        curMaxHealth = .MaxHealth
        curMaxMana = .MaxMana
        curMaxStam = .MaxStam
    End With

    'Mana Regen
    
    Dim iMinMana As Integer, iMinStam As Integer
    iMinMana = GetPercent(curMaxMana, g_Data.MinManaThreshold)
    iMinStam = GetPercent(curMaxStam, g_Data.MinStamThreshold)
    
    'If in Melee-Restam mode, we don't need to keep our mana as high as when we're buffing, so just cut it in half
    If bMeleeRestam Then iMinMana = iMinMana / 2
    
    'If current mana dropped under our minimum mana treshold
    If (curMana <= iMinMana) Then
        
        If bManaRegenMethod = ACT_HEALTH_TO_MANA Then
            bRes = g_Spells.Cast_HealthToMana
            
        Else
        
            'if current mana drops too low
            If curMana <= 40 Then
            
                'check to see if we have enough stam to do an emergency stam2mana
                If curStam >= 40 Then
                    bRes = g_Spells.Cast_Emergency_Stam2Mana

                'we're low on stam, try to get some back!
                Else
                    'TODO: try to see if we can shrug a stam elixir, or a mana potion?
                    bRes = g_Spells.Cast_Emergency_Revitalize
                    
                End If
                
            Else 'we can cast at least one or two more spells with our current mana
            
                'If we're above our minimum stam threshold, we can do a Stam2Mana
                If curStam > iMinStam Then
                    bRes = g_Spells.Cast_StamToMana
                    
                Else 'we first need to get some stamina back
                    bRes = g_Spells.Cast_Revitalize
    
                End If
                
            End If
            
        End If
        
    'Stamina Regen
    ElseIf NeedStamina(bMeleeRestam) Then
        bRes = g_Spells.Cast_Revitalize 'revitalize
    
    'Don't need anything
    Else
        bRes = False
        
    End If
    
Fin:
    CastManaStamRegen = bRes
    Exit Function
ErrorHandler:
    bRes = False
    PrintErrorMessage "clsMacro.CastManaStamRegen() - " & Err.Description
    Resume Fin
End Function

'Returns true if stam potions are enabled and we have stam potions available in inventory
Public Function CanUseStamItem() As Boolean
On Error GoTo ErrorMessage
    Dim bRet As Boolean
    
    If g_ui.Macro.chkUseStamPotion.Checked Then
        ' Look thru the lstStamItems and return first one found
        If Valid(g_ui.Macro.findObjectFromList(g_ui.Macro.lstStamItems)) Then
            bRet = True
        Else
            bRet = (g_Objects.Items.InvFindByName(STR_ITEM_STAM_POTION).Guid <> -1)
        End If
    End If
        
Fin:
    CanUseStamItem = bRet
    Exit Function
ErrorMessage:
    bRet = False
    PrintErrorMessage "Vitals.CanUseStamItem - " & Err.Description
    Resume Fin
End Function

'Find Stam item to use
Public Function findStamItem() As acObject
    g_bFindingItem = True
    Set findStamItem = g_ui.Macro.findObjectFromList(g_ui.Macro.lstStamItems)
    g_bFindingItem = False
End Function

'Find Healing item to use
Public Function findHealItem() As acObject
    g_bFindingItem = True
    Set findHealItem = g_ui.Macro.findObjectFromList(g_ui.Macro.lstHealItems)
    g_bFindingItem = False
End Function

Public Sub idManaStones()
On Error GoTo ErrorMessage
    
    Dim objItem As acObject

    For Each objItem In g_Objects.Items.Inv
        If Valid(objItem) Then
            If (objItem.itemType = ITEM_MANA_STONES) And (objItem.Name Like "*Stone*") Then
                'MyDebug "Vitals: id'ing: " & objItem.Name & " (" & objItem.Mana & ")"
                Call g_Hooks.IDQueueAdd(objItem.Guid)
            End If
        End If
    Next objItem

Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "Error in Vitals.idManaStones: " & Err.Description & " - " & Err.Source
    Exit Sub
End Sub

' Returns the Mana Charge to use if chkUseManaCharge is enabled and
' have Mana Charges in inventory
Public Function findManaCharge() As Boolean
On Error GoTo ErrorMessage
        
    Dim objItem As acObject
    Dim smallestMana As Double
    Dim found As Boolean
    
    Set g_manaItem = Nothing
    Set objItem = Nothing
    
    g_bFindingItem = True

    found = False
    smallestMana = 1000000
    
    'If we use mana stones, loop thru inventory looking for stone with smallest mana charge
    'However, if we don't find a charged one, fallback to useing Mana Charges
    If g_ui.Macro.chkUseManaStone.Checked Then
        
        For Each objItem In g_Objects.Items.Inv
            If Valid(objItem) Then
                If (objItem.itemType = ITEM_MANA_STONES) And (objItem.Mana > 1) And (objItem.Name Like "*Stone*") Then
                    If (objItem.Mana < smallestMana) Then
                        smallestMana = objItem.Mana
                        Set g_manaItem = objItem
                        found = True
                    End If
                End If
            End If
        Next objItem
    
    End If
    
    If found Then GoTo Fin
    
    'Loop through inventory items, looking for Mana Charges
    'However, don't use Massive mana charges As peeps use them as DI's
    For Each objItem In g_Objects.Items.Inv
        
        If Valid(objItem) Then
            If (objItem.Name Like STR_ITEM_MANA_CHARGE) And Not (objItem.Name Like "*Massive*") Then
                Set g_manaItem = objItem
                found = True
                GoTo Fin
            End If
        End If

    Next objItem

Fin:
    findManaCharge = found
    MyDebug "findManaCharge: " & g_manaItem.Guid & " : " & g_manaItem.Name & " (" & g_manaItem.Mana & ")"
    g_bFindingItem = False
    Exit Function
ErrorMessage:
    'PrintErrorMessage "Error in Vitals.findManaCharge: " & Err.Description & " - " & Err.Source
    findManaCharge = False
    g_bFindingItem = False
    Exit Function
End Function

