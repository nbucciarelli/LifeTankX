Attribute VB_Name = "PhatLoot"
Option Explicit

Private m_buyTarget As acObject
Private m_sellTarget As acObject
Private m_Vendor As acObject
Private m_CurStack As acObject
Private m_StartStack As acObject

Private m_iStackSize As Integer
Private m_iCurrentSize As Integer

Private m_vendorInv As colObjects

Private m_buyTargetName As String
Private m_sellTargetName As String

Private m_tmrLagTimer As clsTimer

Private m_numberSold As Long
Private m_totalToSell As Long
Private m_buyCost As Long
Private m_sellCost As Long
Private splitNumber As Integer

Private m_iMode As eSellModes

Private Enum eSellModes
    MODE_SPLIT = 0
    WAIT_SPLIT
    MODE_SELL
    MODE_SELL_CLICK
    WAIT_SELL
    MODE_GIVE_VENDOR
    WAIT_GIVE
    MODE_CLICK_ITEM
    MODE_BUY
    MODE_BUY_CLICK
    WAIT_BUY
    MODE_STACK
    WAIT_STACK
End Enum

'=========================================================
' Setup from the OnApproachVendor event
'=========================================================

Public Sub setVendor(ByVal anObj As acObject)
    Set m_Vendor = anObj
End Sub

Public Sub setVendorInv(ByVal aColl As colObjects)
    Set m_vendorInv = aColl
End Sub

Public Function findInVendorInv(ByVal aGuid As Long) As acObject
    If Valid(m_vendorInv) Then
        If m_vendorInv.Exists(aGuid) Then
            MyDebug "findInVendorInv: Found guid: " & aGuid
            Set findInVendorInv = m_vendorInv.Item(aGuid)
        Else
            MyDebug "findInVendorInv: Not found: aGuid: " & aGuid
            Set findInVendorInv = Nothing
        End If
    Else
        MyDebug "WARNING: m_vendorInv not Valid"
        Set findInVendorInv = Nothing
    End If
End Function

Public Sub vendorEvent(ByVal aSource As String, Optional ByVal obj As acObject)
    If (g_Macro.State <> ST_BUYSELL) Then Exit Sub
    
    If (m_iMode = WAIT_SPLIT And aSource = "OnCreateObject") Then
        If Valid(obj) Then Set m_CurStack = obj
        m_iMode = MODE_GIVE_VENDOR
        'MyDebug "PhatLoot, WAIT_SPLIT: " & aSource
    End If
    
    If (m_iMode = WAIT_GIVE) Then
        m_iMode = MODE_SELL_CLICK
        'MyDebug "PhatLoot, WAIT_GIVE: " & aSource
    End If
    If (m_iMode = WAIT_SELL) Then
        m_iMode = MODE_CLICK_ITEM
        'MyDebug "PhatLoot, WAIT_SELL: " & aSource
    End If
    If (m_iMode = WAIT_BUY) Then
        m_iMode = MODE_STACK
        'MyDebug "PhatLoot, WAIT_BUY: " & aSource
    End If
    If (m_iMode = WAIT_STACK) Then
        m_iMode = MODE_SPLIT
        'MyDebug "PhatLoot, WAIT_STACK: " & aSource
    End If
End Sub

'=========================================================
' Buy and sell Functions
'=========================================================

Public Sub setBuyTarget(ByVal anObj As acObject)
    Set m_buyTarget = anObj
    m_buyTargetName = anObj.Name
End Sub

Public Sub setSellTarget(ByVal anObj As acObject)
    Set m_sellTarget = anObj
    m_sellTargetName = anObj.Name
End Sub

Public Function IsReadyToSell() As Boolean
    If Valid(m_buyTarget) And Valid(m_sellTarget) Then
        IsReadyToSell = True
    Else
        IsReadyToSell = False
    End If
End Function


' The main loop
Public Sub RunState()
On Error GoTo ErrorHandler
    
    If (g_Macro.State <> ST_BUYSELL) Then GoTo Fin
    
    If g_Hooks.VendorID = 0 Then
        PrintMessage "Lost Vendor, exiting"
        Call StopSell
        GoTo Fin
    End If
    
    If Not IsReadyToSell Then
        PrintMessage "Not ready, stoping selling"
        Call StopSell
        GoTo Fin
    End If

    ' Make sure we've waited long enough
    If Not m_tmrLagTimer.Expired Then GoTo Fin
    
    
    ' Ok, lets do something!
    Select Case m_iMode
        Case MODE_SPLIT
        
            If checkPackSpace Then Exit Sub
    
            If (m_numberSold >= m_totalToSell) Then
                ' All done, so cleanup
                PrintMessage "All Done Selling"
                Call StopSell
                GoTo Fin
            End If
  
            ' Split off some number of items to sell
            Dim minPackSpace As Double
            Dim computeDouble As Double
            Dim freePackSpace As Long
            Dim guidCurStack As Long
            Dim objItem As acObject
            
            splitNumber = 0
            
            'MyDebug "MODE_SPLIT"
            
            ' number of slots each item sold takes up
            minPackSpace = CDbl(m_sellCost / 25000)
            freePackSpace = 90 - g_Objects.Items.CountMainInventory
            
            If minPackSpace > 0 Then
                computeDouble = CDbl(freePackSpace / minPackSpace)
            End If
            
            splitNumber = IntRoundDown(computeDouble)
            
            MyDebug "minPackSpace: " & minPackSpace & "  freePackSpace: " & freePackSpace
            MyDebug "splitNumber: " & splitNumber & " computeDouble: " & computeDouble
            
            If (splitNumber < 1) Then
                PrintMessage "You need to free some pack space!"
                Call StopSell
                GoTo Fin
            End If
            
            If (m_iCurrentSize >= m_iStackSize) Then
                'We need to find the next stack of items to sell
                If (g_ui.Main.chkSellAll.Checked) Then
                    PrintMessage "Searching for next stack of " & m_sellTargetName
                    'MyDebug "PhatLoot: Searching for other items named: " & m_sellTargetName
                    For Each objItem In g_Objects.Items.Inv
                        'Same name and in Main pack
                        If (objItem.Name = m_sellTargetName) And (objItem.Container = g_Objects.Player.Guid) Then
                            Set m_StartStack = objItem
                            Set m_sellTarget = objItem
                            m_iStackSize = m_StartStack.stackCount
                            m_iCurrentSize = 0
                            GoTo Done
                        End If
                    Next objItem
Done:
                Else
                    PrintMessage "All Done Selling"
                    Call StopSell
                    GoTo Fin
                End If
            End If
            
            Call g_Hooks.SelectItem(m_StartStack.Guid)
            
            If (splitNumber > m_StartStack.stackCount) Then
                splitNumber = m_StartStack.stackCount
                Set m_CurStack = m_StartStack
                m_iMode = MODE_GIVE_VENDOR
            Else
                'SetStackCount
                g_Hooks.SelectedStackCount = splitNumber
                Call g_Hooks.SelectItem(m_StartStack.Guid)
                Call g_Hooks.MoveItem(g_Hooks.CurrentSelection, g_Objects.Player.Guid, 0, False)
                m_iMode = WAIT_SPLIT
            End If
            
            'MyDebug "m_StartStack: " & m_StartStack.Guid
            
            Call m_tmrLagTimer.SetNextTime(1)    ' 2 seconds to split stack
            
        Case WAIT_SPLIT
            Call m_tmrLagTimer.SetNextTime(1)
        
        Case MODE_GIVE_VENDOR
            
            If checkPackSpace Then Exit Sub
            
            ' Take the split off items and put in vendor window
            
            'MyDebug "MODE_GIVE_VENDOR"
            
            If Valid(m_CurStack) Then
                MyDebug "Giving: " & m_CurStack.Guid & " to: " & g_Hooks.VendorID
                Call g_Hooks.GiveItem(m_CurStack.Guid, g_Hooks.VendorID)
            Else
                PrintMessage "Failed to sell item stack to vendor"
                Call StopSell
                GoTo Fin
            End If
            
            m_iMode = MODE_SELL_CLICK
            Call m_tmrLagTimer.SetNextTime(1)    ' 2 seconds to give vendor
        
        Case MODE_SELL_CLICK
        
            If checkPackSpace Then Exit Sub
            
            'MyDebug "MODE_SELL_CLICK"
            
            ' call mouse movement and button pressing
            Utils.ClickVendorButton
            
            ' make sure to update the number sold
            m_numberSold = m_numberSold + splitNumber
            m_iCurrentSize = m_iCurrentSize + splitNumber
            
            MyDebug "m_numberSold: " & m_numberSold
            
            m_iMode = WAIT_SELL
            Call m_tmrLagTimer.SetNextTime(1)
            
        Case WAIT_SELL
            Call m_tmrLagTimer.SetNextTime(1)
            
        '-----------------------------------------------------
        ' Now start the Buy part of the routine
        '-----------------------------------------------------
        Case MODE_CLICK_ITEM
        
            'MyDebug "MODE_CLICK_ITEM"
            
            ' Make sure we are on the Item tab of the vendor
            Utils.ClickVendorItemsTab
            
            m_iMode = MODE_BUY
            Call m_tmrLagTimer.SetNextTime(1)    ' 1 seconds to switch to item tab
        
        Case MODE_BUY
        
            'MyDebug "MODE_BUY"
                        
            ' How many to buy?
            Dim buyNumber As Integer
            Dim guidBuyTarget As Long
            Dim totalCash As Long
                  
            'm_buyCost = m_buyTarget.Value
            
            If m_buyCost <= 0 Then m_buyCost = 287500
            buyNumber = Utils.IntRoundDown(g_ds.XpTracker.TotalPyreals / m_buyCost)
            
            MyDebug "buyNumber: " & buyNumber & "  totalCash: " & g_ds.XpTracker.TotalPyreals
            
            'MyDebug "m_buyTarget.guid: " & m_buyTarget.Guid
            
            If (buyNumber = 0) Then
                'We don't have enough cash, so reset back to selling part
                m_iMode = MODE_SPLIT
                PrintMessage "Not enough to buy even one, back to selling"
                GoTo Fin
            End If
            
            Call g_Hooks.SelectItem(m_buyTarget.Guid)
            g_Hooks.SelectedStackCount = buyNumber
            
            'MyDebug "after m_buyTarget.guid: " & m_buyTarget.Guid
            
            m_iMode = MODE_BUY_CLICK
            Call m_tmrLagTimer.SetNextTime(1)    ' 2 seconds to select item
            
        Case MODE_BUY_CLICK
            
            'MyDebug "MODE_BUY_CLICK"
            
            ' Click the Buy button on vendor screen
            Utils.ClickVendorButton
            
            m_iMode = WAIT_BUY
            Call m_tmrLagTimer.SetNextTime(1)
            
        Case WAIT_BUY
            Call m_tmrLagTimer.SetNextTime(1)
        
        Case MODE_STACK
            ' Stack up all the items we just bought
            
            'MyDebug "MODE_STACK"
            
            For Each objItem In g_Objects.Items.Inv
                If (objItem.Name = m_buyTargetName) And (objItem.stackCount < objItem.StackMax) And (objItem.Container = g_Objects.Player.Guid) Then
                    Call g_Hooks.SelectItem(objItem.Guid)
                    Call g_Hooks.MoveItem(g_Hooks.CurrentSelection, g_Objects.Player.Guid, 0, True)
                End If
            Next objItem

            m_iMode = MODE_SPLIT
            Call m_tmrLagTimer.SetNextTime(1)    ' 5 seconds to stack items
        
        Case Else
            ' Bad State! ERROR!
            PrintErrorMessage "Invalid state in PhatLoot Runstate, exiting"
            MyDebug "Invalide m_iMode: " & m_iMode
            Call g_Macro.GoIdle
    End Select

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PhatLoot.RunState - " & Err.Description
    Call g_Macro.GoIdle
    Resume Fin
End Sub

Private Function checkPackSpace() As Boolean
    ' Check to see if backpack is full
    If g_Objects.Items.BackpackFull Then
        PrintMessage "Main Pack is full, clear some space then start over"
        Call StopSell
        checkPackSpace = True
    Else
        checkPackSpace = False
    End If

End Function

Public Sub StopSell()
On Error GoTo ErrorHandler
        
        ' All done, so cleanup
        m_totalToSell = 0
        m_numberSold = 0
        m_buyCost = 0
        m_sellCost = 0
        Set m_sellTarget = Nothing
        g_ui.Main.lblSellTarget.Text = "None"
        Call m_tmrLagTimer.Reset
        m_tmrLagTimer.Enabled = False
        Call g_Macro.GoIdle
        m_iMode = -1
        g_ui.Main.btnStartSell.Text = "Start"

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PhatLoot.StopSell - " & Err.Description
    Exit Sub
End Sub

Public Sub StartSell()
On Error GoTo ErrorHandler

    PrintMessage "In StartSell"

    Dim objItem As acObject
    Dim stackCount As Long
    
    If g_Hooks.VendorID = 0 Then
        PrintMessage "You must be at a vendor first"
        GoTo Fin
    End If
    
    m_totalToSell = 0
    m_numberSold = 0
    
    stackCount = m_sellTarget.stackCount

    If Valid(m_Vendor) And m_Vendor.VendorFractBuy > 0 And stackCount > 0 Then
        MyDebug "vendor FractBuy: " & m_Vendor.VendorFractBuy
        m_sellCost = (m_sellTarget.Value / m_Vendor.VendorFractBuy) / stackCount
    ElseIf stackCount > 0 Then
        m_sellCost = m_sellTarget.Value / stackCount
    End If
    
    If Valid(m_Vendor) Then
        MyDebug "vendor FractSell: " & m_Vendor.VendorFractSell
        m_buyCost = (m_buyTarget.Value * m_Vendor.VendorFractSell)
    Else
        m_buyCost = m_buyTarget.Value
    End If
    
    If InStr(m_buyTarget.Name, "Trade Note") Then m_buyCost = m_buyTarget.Value * 1.18

    MyDebug "m_sellCost: " & m_sellCost & "  m_buyCost: " & m_buyCost

    Set m_StartStack = m_sellTarget
    Set m_CurStack = m_StartStack
    
    m_iStackSize = m_StartStack.stackCount
    m_iCurrentSize = 0
    
    If (g_ui.Main.chkSellAll.Checked) Then
        MyDebug "PhatLoot: Searching for other items named: " & m_sellTarget.Name
        For Each objItem In g_Objects.Items.Inv
            ' Same name and in Main pack
            If (objItem.Name = m_sellTarget.Name) And (objItem.Container = g_Objects.Player.Guid) Then
                m_totalToSell = m_totalToSell + objItem.stackCount
            End If
        Next objItem
    Else
        m_totalToSell = stackCount
    End If
        
    MyDebug "PhatLoot.StartSell"
    
    If Not IsReadyToSell() Then
        PrintMessage "Not Ready to sell, internal variables not set correctly"
    End If
    
    'MyDebug "m_buyTarget: " & m_buyTarget.Name & " m_sellTarget: " & m_sellTarget.Name
    'MyDebug "m_sellCost: " & m_sellCost & "  m_buyCost: " & m_buyCost
    'MyDebug "m_totalToSell: " & m_totalToSell
    
    m_iMode = MODE_SPLIT
    
    Call g_Macro.SetState(ST_BUYSELL)
    
    Set m_tmrLagTimer = CreateTimer
    m_tmrLagTimer.Enabled = True
    m_tmrLagTimer.SetNextTime (1)
    
    g_ui.Main.btnStartSell.Text = "STOP"
  
Fin:
    'If g_Macro.Paused Then Call TogglePause
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PhatLoot.StartSell - " & Err.Description
    Resume Fin
End Sub


'=========================================================

Public Function PassActiveFilters(objItem As acObject, _
                                    Optional ByVal bSalvageFilters As Boolean = True, _
                                    Optional ByVal bArmorFilters As Boolean = True, _
                                    Optional ByVal bWeaponFilters As Boolean = True, _
                                    Optional ByVal bWandFilters As Boolean = True, _
                                    Optional ByVal bMajorMinorFilters As Boolean = True, _
                                    Optional ByVal bValueFilter As Boolean = False, _
                                    Optional ByVal bHighManaFilter As Boolean = True)
On Error GoTo ErrorHandler

    'Assume true
    PassActiveFilters = True
    
    If Not Valid(objItem) Then
        PrintErrorMessage "PhatLoot.PassActiveFilters - invalid objItem !"
        GoTo Fin
    End If
    
    'MyDebug "PhatLoot.PassActiveFilters for: " & objItem.Name
    
    If objItem.IsRare Then
        PassActiveFilters = True
        GoTo Fin
    End If

    If bMajorMinorFilters And g_ui.Loot.chkLootAny.Checked Then
        Dim vSpellName As Variant
        Dim bPass As Boolean
            
        'Majors
        If g_ui.Loot.chkPickupMajors.Checked And (objItem.HasMajors Or objItem.HasEpics) Then
            bPass = True
            If g_ui.Loot.chkMajorIgnoreBane.Checked And Valid(objItem.Spells) Then
                bPass = False
                'In case the item has 2 majors...
                For Each vSpellName In objItem.Spells
                    If InStr(vSpellName, "Major") Or InStr(vSpellName, "Epic") Then
                        If InStr(vSpellName, " Bane") < 1 Then
                            bPass = True
                            Exit For
                        Else
                            MyDebug "... Major Bane Detected : " & vSpellName
                        End If
                    End If
                Next vSpellName
            End If
            If bPass Then GoTo Fin
        End If
        
        'Minors
        If g_ui.Loot.chkPickupMinors.Checked And objItem.HasMinors Then
            Dim bWentInSubRule As Boolean
            
            bPass = True
            bWentInSubRule = False
            
            'Minor Wards
            If g_ui.Loot.chkMinorWard.Checked Then
                bWentInSubRule = True
                bPass = False
                For Each vSpellName In objItem.Spells
                    If InStr(vSpellName, "Minor") And InStr(vSpellName, "Ward") Then
                        bPass = True
                        Exit For
                    End If
                Next vSpellName
            End If
            
            'Minor Attribute
            If g_ui.Loot.chkMinorAttribute.Checked And ((bWentInSubRule And Not bPass) Or Not bWentInSubRule) Then
                bPass = False
                bWentInSubRule = True
                For Each vSpellName In objItem.Spells
                    If InStr(vSpellName, "Minor") Then
                        If InStr(vSpellName, "Strenght") _
                        Or InStr(vSpellName, "Endurance") _
                        Or InStr(vSpellName, "Coordination") _
                        Or InStr(vSpellName, "Quickness") _
                        Or InStr(vSpellName, "Focus") _
                        Or InStr(vSpellName, "Willpower") Then
                            bPass = True
                            Exit For
                        End If
                    End If
                Next vSpellName
            End If
            
            'Minor Masteries
            If g_ui.Loot.chkMinorMastery.Checked And ((bWentInSubRule And Not bPass) Or Not bWentInSubRule) Then
                bPass = False
                bWentInSubRule = True
                For Each vSpellName In objItem.Spells
                    If InStr(vSpellName, "Minor") And InStr(vSpellName, "Mastery") Then
                        bPass = True
                        Exit For
                    End If
                Next vSpellName
            End If
            
            If bPass Then GoTo Fin
        End If
    End If
    
    If bValueFilter And g_ui.Loot.chkPickupValuable.Checked Then
        If objItem.Value >= g_Data.LootMinValue Then
            'apply Value/Burden filter
            If g_ui.Loot.chkBurdenRatio.Checked And (objItem.Burden <> 0) And (objItem.Burden <> -1) Then
                Dim ratio As Double
                ratio = Round(objItem.Value / objItem.Burden)
                If ratio >= g_Data.LootMinBurdenRatio Then GoTo Fin
            Else    'no burden check to do
                GoTo Fin
            End If
        End If
    End If
            
    'Armor Filters
    If bArmorFilters And g_ui.Loot.chkLootArmors.Checked And (objItem.itemType = ITEM_ARMOR) Then
        If g_Data.LootFilters.PassFilters(objItem, g_Data.LootFilters.ArmorFilters) Then GoTo Fin
    End If
    
    'Weapon Filters
    If bWeaponFilters And g_ui.Loot.chkLootWeapons.Checked And ((objItem.itemType = ITEM_MELEE_WEAPON) Or (objItem.itemType = ITEM_MISSILE_WEAPON)) Then
        If g_Data.LootFilters.PassFilters(objItem, g_Data.LootFilters.WeaponFilters) Then GoTo Fin
    End If
    
    'Wand Filters
    If bWandFilters And g_ui.Loot.chkLootWands.Checked And (objItem.itemType = ITEM_WAND) Then
        If g_Data.LootFilters.PassFilters(objItem, g_Data.LootFilters.WandFilters) Then GoTo Fin
    End If
    
    'Salvage Filters
    If bSalvageFilters And g_ui.Loot.chkLootSalvages.Checked Then
        If g_Data.LootFilters.PassFilters(objItem, g_Data.LootFilters.SalvageFilters) Then GoTo Fin
    End If
    
    'Unknown scrolls
    If g_ui.Loot.chkUnknownScrolls.Checked And (InStr(LCase(objItem.Name), "scroll") Or (objItem.itemType = ITEM_SCROLL)) Then
        Dim sName As String
        Dim SpellID As Long
        Dim aSkill As Long
        Dim aSpell As clsSpell
        'get scroll name
        sName = objItem.Name
        'get spell ID
        SpellID = objItem.AssociatedSpellId
        
        MyDebug "PhatLoot.PassLootFilter scroll: " & sName & " spellId: " & SpellID
        
        'see if we know it
1        If Not g_Filters.SpellLearned(SpellID) Then
            'Make sure it's a school we have trained
2            Set aSpell = g_Spells.FindSpellByID(SpellID)
3            If Valid(aSpell) Then
4                MyDebug "PhatLoot.scrolls: Name: " & sName & "   school: " & aSpell.SpellSchool & " isTrained: " & g_Hooks.Skill(aSpell.SpellSchool)
5                aSkill = MagicSchoolToSkillId(aSpell.SpellSchool)
6                If g_Hooks.SkillTrainLevel(aSkill) <> eUntrained Then
                    ' add to array so we can read it later
                    MyDebug "Looting unknown scroll: " & sName
                    ' all done, pick up!
                    GoTo Fin
                End If
'5                If (g_Hooks.Skill(aSpell.SpellSchool) > 0) Then
'                    ' add to array so we can read it later
'6                    MyDebug "Looting unknown scroll: " & sName
'                    ' all done, pick up!
'                    GoTo Fin
'                End If
            Else
                MyDebug "Looting unknown scroll: " & sName
                GoTo Fin
            End If
        End If
    End If
    
    'Should we look at items Mana?
    If g_ui.Macro.chkRechargeManaStones.Checked And g_Macro.Loot.canLootHighManaItem And bHighManaFilter Then
        If objItem.Mana > g_Data.HighManaValue Then
            MyDebug "Looting High Mana Item: " & objItem.Name & " mana: " & objItem.Mana
            GoTo Fin
        'Else
        '    locDebug "PhatLoot: hasEmptyManaStone, but not enough mana: " & objItem.Name & " mana: " & objItem.Mana
        End If
    End If

    'MyDebug objItem.Name & " failed to pass any filter"
    'Didn't pass any filter
    PassActiveFilters = False
    
Fin:
    Exit Function
ErrorHandler:
    PassActiveFilters = True    'better to loot/keep it than risk destroying it if an error happens here
    PrintErrorMessage "PassActiveFilters - " & Err.Description & "  line: " & Erl
    Resume Fin
End Function

Public Function IsImportantItem(objItem As acObject, Optional ByVal bMainPackOnly As Boolean = True)
On Error GoTo ErrorHandler

    IsImportantItem = True
    
    If Not Valid(objItem) Then
        PrintErrorMessage "PhatLoot.IsImportantItem : invalid objItem - ignoring"
        GoTo Fin
        
    'Must be in main inventory
    ElseIf bMainPackOnly And (objItem.Container <> g_Objects.Player.Guid) Then GoTo Fin
    
    'Must not be wielded
    ElseIf objItem.Wielder = g_Objects.Player.Guid Then GoTo Fin
        
    'Must not be equiped
    ElseIf objItem.Equiped Then GoTo Fin
    
    'Check item type
    ElseIf objItem.itemType = ITEM_UNKNOWN _
        Or objItem.itemType = ITEM_TRADENOTE _
        Or objItem.itemType = ITEM_CONTAINER Then GoTo Fin
        
    'Must not be in exception list
    ElseIf g_Data.Exceptions.Items.Exists(objItem.Guid) Then GoTo Fin
    
    'Must not have tinks on it
    ElseIf (objItem.TinkCount > 0) Then GoTo Fin
    
    'Else it's not important
    Else
        IsImportantItem = False
    End If
        
Fin:
    Exit Function
ErrorHandler:
    IsImportantItem = True
    PrintErrorMessage "IsImportantItem - " & Err.Description
    Resume Fin
End Function


Public Sub SellAtVendor()
On Error GoTo ErrorHandler

    Dim objItem As acObject
    
    If g_Hooks.VendorID = 0 Then
        PrintMessage "You must be at a vendor first"
        GoTo Fin
    End If
        
    MyDebug "Selling items at vendor..."
    
    For Each objItem In g_Objects.Items.Inv
            
        If IsImportantItem(objItem) Then GoTo NextItem
        
        'Skip other special case items
        Select Case objItem.itemType
            Case ITEM_CONTAINER, ITEM_SALVAGE, ITEM_TRADENOTE, ITEM_UNKNOWN
                GoTo NextItem
        End Select
        
        'Must not be a good item
        If PassActiveFilters(objItem) Then GoTo NextItem
        
        'Put it in the sell window
        Call g_Hooks.GiveItem(objItem.Guid, g_Hooks.VendorID)
        
NextItem:
    Next objItem

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "SellAtVendor - " & Err.Description
    Resume Fin
End Sub

Public Sub PutSelectionInUst()
On Error GoTo ErrorHandler
    
    If g_Hooks.CurrentSelection = 0 Then
        PrintErrorMessage "Please select a valid item"
    Else
        Call PutSalvageInUst(g_Objects.FindObject(g_Hooks.CurrentSelection).MaterialType)
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PutSelectionInUst - " & Err.Description
    Resume Fin
End Sub

Public Sub PutSalvageInUst(ByVal lSalvageType As Long)
On Error GoTo ErrorHandler

    Dim objItem As acObject
    MyDebug "Putting salvages in ust..."
    
    For Each objItem In g_Objects.Items.Inv
        
        If Not Valid(objItem) Then GoTo NextItem
        
        'Must be of same salvage kind
        If objItem.MaterialType <> lSalvageType Then GoTo NextItem
        
        'Must have a workmanship
        If objItem.Workmanship < 1 Then GoTo NextItem
        
        'Must not be an important item
        If IsImportantItem(objItem) Then GoTo NextItem
        
        'Skip special items
        Select Case objItem.itemType
            'Case ITEM_CONTAINER, ITEM_SALVAGE, ITEM_SCROLL, ITEM_TRADENOTE, ITEM_UNKNOWN
            Case ITEM_CONTAINER, ITEM_SCROLL, ITEM_TRADENOTE, ITEM_UNKNOWN
                GoTo NextItem
        End Select
        
        'Check all filters but salvage filters
        If PassActiveFilters(objItem, False) Then GoTo NextItem
        
        'Add to ust
        Call g_Hooks.SalvagePanelAdd(objItem.Guid)
        
NextItem:
    Next objItem

Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "PutSalvageInUst - " & Err.Description
    Resume Fin
End Sub

' Check to see if this object is Valid Salvage
Public Function CheckHighManaItem(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    'MyDebug "PhatLoot.CheckHighManaItem: " & objItem.Name
    
        If PhatLoot.IsImportantItem(objItem) Then
            bRet = False
        ElseIf objItem.HasMajors Then
            bRet = False
        ElseIf objItem.IsRare Then
            bRet = False
        ElseIf PassActiveFilters(objItem, False, True, True, True, True, False, False) Then
            'Passed some other filter
            bRet = False
        ElseIf objItem.Mana < g_Data.HighManaValue Then
            bRet = False
        Else
            'Ok we found a HighMana item, add it to our list
            bRet = True
        End If

Fin:
    CheckHighManaItem = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "PhatLoot.CheckHighManaItem - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function IsWorthAssessing(lItemType As Long) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    
    Select Case lItemType
        'Items that don't give any extra useful information when IDed
        Case ITEM_PYREAL, _
            ITEM_UST, _
            ITEM_CONTAINER, _
            ITEM_SALVAGE, _
            ITEM_TRADENOTE, _
            ITEM_FOOD, _
            ITEM_ARROW, _
            ITEM_BUNDLE, _
            ITEM_COMPS, _
            ITEM_HEALING_KIT, _
            ITEM_LOCKPICK
            
            bRet = False
            
        Case Else
            bRet = True
    End Select
    
Fin:
    IsWorthAssessing = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "PhatLoot.IsWorthAssessing - " & Err.Description
    Resume Fin
End Function

Public Function doAutoStacking() As Boolean
On Error GoTo ErrorHandler

    Dim colInv As colObjects
    Dim bRet As Boolean
    Dim objItem1 As acObject
    Dim objItem2 As acObject
    Dim bottomCount As Integer
    Dim topCount As Integer
    
    bRet = False
    Set colInv = New colObjects
    
    For Each objItem1 In g_Objects.Items.Inv
        If (objItem1.stackCount < objItem1.StackMax) And (objItem1.Container = g_Objects.Player.Guid) Then
            Call colInv.addObject(objItem1)
        End If
    Next objItem1
    
    If colInv.Count > 1 Then
    
        Dim itemList() As acObject
        
        itemList = ColToArray(colInv)
        
        For topCount = UBound(itemList) To LBound(itemList) Step -1
            Set objItem1 = itemList(topCount)
            If (objItem1.stackCount < objItem1.StackMax) Then
                For bottomCount = LBound(itemList) To UBound(itemList)
                    Set objItem2 = itemList(bottomCount)
                    If (objItem2.Name = objItem1.Name) And (objItem2.stackCount < objItem2.StackMax) And (objItem1.Guid <> objItem2.Guid) Then
                        Call g_Hooks.MoveItemEx(objItem1.Guid, objItem2.Guid)
                        bRet = True
                        GoTo Fin
                    End If
                Next bottomCount
            End If
        Next topCount

    End If
    
Fin:
    Set colInv = Nothing
    Set objItem1 = Nothing
    Set objItem2 = Nothing
    doAutoStacking = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "PhatLoot.doAutoStacking - " & Err.Description
    Resume Fin
End Function
