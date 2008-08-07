VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSalvager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Salvager                                          [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False

Private Enum eSubStates
    ST_NONE
    ST_PREPARE_UST
    ST_LOAD_SALVAGES
    ST_DO_SALVAGE
    ST_WAIT_FOR_COMPLETION
    ST_DONE
End Enum

Private Const TRANSITION_DELAY = 0.3    'seconds
Private Const SECURE_TIME = 6           'seconds

Private Const MAX_UNITS_PER_BAG = 100
Private Const BAGS_MERGE_TOLERANCE = 6      'allow the loss of max X units of salvages when merging two bags together
Private Const BAGS_COMPLETION_TOLERANCE = 6 'allow the loss of max X units when completing the bag with an item

Private WithEvents m_tmrTransition As clsTimer
Attribute m_tmrTransition.VB_VarHelpID = -1
Private WithEvents m_tmrSecure As clsTimer
Attribute m_tmrSecure.VB_VarHelpID = -1
Private m_tmrNextTime As clsTimer

Private m_State As eSubStates
Private m_objUst As acObject
Private m_bComplete As Boolean  'tell if salvaging was complete or interupted

Private m_IdQueue As colObjects
Private m_ItemsToSalvage As colObjects

Private m_LootedItemsToSalvage As colObjects 'All the looted items we want to salvage

Private m_colSalvages As colObjects 'List of all the items from the main pack we can salvage
Private m_colBatch As colObjects
Private m_batchState As Integer
Private m_colItemsInUse As colObjects  'List of all the items that are still being usted by AC and haven't been removed yet

Public Event OnSalvageStopped(ByVal bComplete As Boolean)

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    m_State = ST_NONE
    Set m_tmrTransition = CreateTimer
    Set m_tmrSecure = CreateTimer
    Set m_tmrNextTime = CreateTimer
    Set m_IdQueue = New colObjects
    Set m_ItemsToSalvage = New colObjects
    Set m_colBatch = New colObjects
    Set m_colItemsInUse = New colObjects
    Set m_LootedItemsToSalvage = New colObjects
    Set m_objUst = Nothing
    Call Reset
End Sub

Public Sub Reset()
    m_bComplete = False
    Call m_tmrSecure.Reset
    Call m_tmrTransition.Reset
    Call m_tmrNextTime.Reset
    Call SetState(ST_NONE)
End Sub

Private Sub Class_Terminate()
    Set m_tmrTransition = Nothing
    Set m_tmrSecure = Nothing
    Set m_tmrNextTime = Nothing
    Set m_objUst = Nothing
    Set m_ItemsToSalvage = Nothing
    Set m_IdQueue = Nothing
    Set m_colSalvages = Nothing
    Set m_colBatch = Nothing
    Set m_colItemsInUse = Nothing
    Set m_LootedItemsToSalvage = Nothing
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                            Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Public Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function AddLootToSalvageList(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    bRet = m_LootedItemsToSalvage.addObject(objItem)
    MyDebug "AddLootToSalvageList: " & bRet & " :: " & objItem.Name

Fin:
    AddLootToSalvageList = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSalvager.AddLootToSalvageList - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function UstAvailable() As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not Valid(m_objUst) Then
        Set m_objUst = g_Objects.Items.InvFindByName(STR_ITEM_UST, False)
    End If
    
    bRet = Valid(m_objUst)
    
Fin:
    UstAvailable = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSalvager.UstAvailable - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function StartSalvaging() As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Not UstAvailable Then
        PrintMessage "Trying to salvage, but no Ust available!"
        MyDebug "clsSalvager.StartSalvaging : no ust available"
        GoTo Fin
    End If
    
    If MakeAllBatches(m_colBatch) Then
        m_bComplete = False
        Call SetState(ST_PREPARE_UST)
        Call m_tmrSecure.SetNextTime(3)
        Call m_tmrNextTime.ExpireNow
        bRet = True
    End If
    
    locDebug "StartSalvaging called, bRet = " & bRet
    
Fin:
    StartSalvaging = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSalvager.StartSalvaging - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Sub RunState()
On Error GoTo ErrorHandler
   
   Dim bValidBatch As Boolean
        
   If m_tmrTransition.Enabled Then GoTo Fin
    
    'FIXME: put secure timer back
    If m_tmrSecure.Expired Then
        PrintWarning "clsSalvage.RunState[" & GetStateString & "] - Secure Timer expired : exiting state."
        Call StopSalvage("RunState - Secure Timer Expired")
        GoTo Fin
    End If

    If g_Macro.NeedHealing("clsLoot.RunState") Then
        Call StopSalvage("RunState - Need Healing")
        GoTo Fin
    End If

    If m_tmrNextTime.Expired Then
        Select Case m_State
                       
            Case ST_PREPARE_UST
                
                If Not OpenUst Then
                    PrintErrorMessage "Salvager: unable to open ust, exiting"
                    Call StopSalvage("Unable to open Ust")
                    GoTo Fin
                End If
                
                'Wait half a second before putting items on the ust
                Call m_tmrNextTime.SetNextTime(0.5)
                Call SetState(ST_LOAD_SALVAGES)
                Call m_tmrSecure.SetNextTime(3)
                
            Case ST_LOAD_SALVAGES
                If Not Valid(m_colBatch) Then
                    bValidBatch = MakeAllBatches(m_colBatch)
                ElseIf m_colBatch.Count < 1 Then
                    bValidBatch = MakeAllBatches(m_colBatch)
                Else
                    bValidBatch = True
                End If
                
                If bValidBatch Then
                    Dim objItem As acObject
                    
                    For Each objItem In m_colBatch
                        If (objItem.UserData(INT_SALVAGECOUNT) > 2) Then
                            MyDebug "Salvager: Found bad salvage item: " & objItem.Guid & " : " & objItem.Name
                            If m_colBatch.Remove(objItem.Guid) Then
                                MyDebug "clsSalvager - Removing bad salvage " & objItem.Name & " from batch"
                            End If
                        Else
                            'Set the UST_LOADED count on each item
                            Dim iCount As Integer
                            iCount = objItem.UserData(INT_SALVAGECOUNT) + 1
                            Call objItem.SetUserData(INT_SALVAGECOUNT, iCount)
                            
                            locDebug "Salvager: Adding " & objItem.Name & " to ust..."
                            Call g_Service.AddToUst(objItem)
                        End If
                    Next objItem
                    
                    'Wait 0.5s before pushing the salvage button
                    Call m_tmrNextTime.SetNextTime(0.5)
                    Call SetState(ST_DO_SALVAGE)
                    Call m_tmrSecure.SetNextTime(4)
                Else
                    Call SetState(ST_DONE)
                End If
                
            'Hit the salvage button
            Case ST_DO_SALVAGE
                'PrintMessage "ST_DO_SALVAGE - FIXME: Push the salvage button here!"
                Call SetState(ST_WAIT_FOR_COMPLETION)
                Call Utils.ClickSalvageButton
                
                'Wait 0.5s before adding new salvages to the ust
                Call m_tmrNextTime.SetNextTime(0.5)
                
            Case ST_WAIT_FOR_COMPLETION
                'Do nothing here, we're just waiting for the "You salvage XXX for XXX" green message
                
            Case ST_DONE
                'We're done salvaging, leave salvage state
                m_bComplete = True
                If Not g_ui.Options.chkFilterSalvageMsg.Checked Then
                    PrintMessage "Auto-Salvaging Complete!"
                End If
                Call StopSalvage("RunState - ST_DONE")
                
            Case Else
                PrintWarning "clsSalvager.RunState - Invalid State #" & m_State
                Call StopSalvage("RunState - Invalid State #" & m_State)
                
        End Select
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.RunState - " & Err.Description
    Resume Fin
End Sub

Public Sub StopSalvage(Optional ByVal sSrc As String = "None")
    locDebug "StopSalvage - From : " & sSrc
    Call SetState(ST_NONE)
    Call m_tmrTransition.SetNextTime(TRANSITION_DELAY)
End Sub

'Prints out the current salvages list
Public Sub GetSalvagesList()
On Error GoTo ErrorHandler

    Dim objItem As acObject

    'Loop through inventory items, looking for valid salvages
    For Each objItem In m_LootedItemsToSalvage
        If Valid(objItem) Then
            PrintMessage "Salvage: " & objItem.Name
        End If
    Next objItem

Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.GetSalvagesList - " & Err.Description
    Resume Fin
End Sub

' Check to see if this object is Valid Salvage
Public Function CheckValidSalvageItem(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    locDebug "clsSalvager: CheckValidSalvageItem: " & objItem.Name
    
        'Check for bad salvage item
        If (objItem.UserData(INT_SALVAGECOUNT) > 2) Then
            bRet = False
        ElseIf objItem.Workmanship < 1 Then
            bRet = False
        ElseIf objItem.itemType = ITEM_SALVAGE Then
            bRet = False
        ElseIf PhatLoot.IsImportantItem(objItem) Then
            bRet = False
        ElseIf objItem.HasMajors Then
            bRet = False
        ElseIf IsWorthAssessing(objItem.itemType) And (objItem.LastIdTime = 0) Then
            'If item hasn't been assessed return
            bRet = False
        ElseIf PassActiveFilters(objItem, False) Then
            'Passed some other filter
            bRet = False
        Else
            'Ok we found a salvageable item, add it to our list
            bRet = True
        End If

Fin:
    CheckValidSalvageItem = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSalvager.CheckValidSalvageItem - " & Err.Description
    bRet = False
    Resume Fin
End Function

'List all the valid non-salvaged-yet items in main inventory we can salvage
Public Sub MakeValidSalvagesList()
On Error GoTo ErrorHandler
    
    Dim objItem As acObject
    
    'Start a new collection
    Set m_colSalvages = New colObjects
    
    locDebug "clsSalvager: MakeValidSalvagesList: Items.Inv: " & g_Objects.Items.Inv.Count
    
    'Loop through inventory items, looking for valid salvages
    For Each objItem In g_Objects.Items.Inv
        
        locDebug "MVSL: Testing: " & objItem.Name
        
        'Check for bad salvage item
        If (objItem.UserData(INT_SALVAGECOUNT) > 2) Then
            locDebug "MVSL: INT_SALVAGECOUNT: " & objItem.UserData(INT_SALVAGECOUNT)
            GoTo NextItem
        End If
        
        'Must have a workmanship
        If objItem.Workmanship < 1 Then
            locDebug "MVSL: Workmanship: " & objItem.Workmanship
            GoTo NextItem
        End If
        
        'Must not be a bag of salvage
        If objItem.itemType = ITEM_SALVAGE Then
            locDebug "MVSL: ITEM_SALVAGE true"
            GoTo NextItem
        End If

        'Must not be an important item
        If PhatLoot.IsImportantItem(objItem) Then
            locDebug "MVSL: PhatLoot IsImportantItem"
            GoTo NextItem
        End If
        
        'Check all filters but salvage filters
        If PhatLoot.PassActiveFilters(objItem, False) Then
            locDebug "MVSL: PhatLoot PassActiveFilters"
            GoTo NextItem
        End If
        
        'Salvage must be in our "salvages to pickup" list, even if the "loot salvage" is unchecked
        If Not g_Data.LootFilters.PassFilters(objItem, g_Data.LootFilters.SalvageFilters) Then
            locDebug "MVSL: NOT in SalvageFilter list"
            GoTo NextItem
        End If
        
        'Never Salvage a major (yeah, I know major banes should be salvage :)
        If objItem.HasMajors Then
            locDebug "MVSL: HasMajors: " & objItem.HasMajors
            GoTo NextItem
        End If
        
        'If item hasn't been assessed return
        If IsWorthAssessing(objItem.itemType) And (objItem.LastIdTime = 0) Then
            locDebug "IsWorthAssessing, LastIdTime = " & objItem.LastIdTime
            'Call g_Hooks.IDQueueAdd(objItem.Guid)
            GoTo NextItem         'FIX: 11/11/2004 - Prevent non assessed items from being salvaged
        End If
        
        'Ok we found a salvageable item, add it to our list
        If Not m_colSalvages.addObject(objItem) Then
            PrintWarning "clsSalvager.MakeValidSalvagesList - Unable to add item " & objItem.Name
        Else
            locDebug "clsSalvager.MakeValidSalvagesList - Added: " & objItem.Name
        End If
        
NextItem:
    Next objItem
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.MakeValidSalvagesList - " & Err.Description
    Resume Fin
End Sub

'Get a list of all the partial bags of salvage of the first valid type we can find in inventory
'and return the collection, optionally giving the bag with the most units on it
'If lSalvateType is not specified, the function will pick the first bags with a valid salvage kind
'FIXME: make private
Public Function FindPartialBags(ByRef colBags As colObjects, _
                                    Optional ByVal lSalvageType As Long = -1, _
                                    Optional ByVal colIgnoreList As colObjects, _
                                    Optional ByRef objFullestBagOut As acObject, _
                                    Optional ByVal iMinWork As Integer = 1, _
                                    Optional ByVal iMaxWork As Integer = 10) As Boolean
On Error GoTo ErrorHandler

    Dim objItem As acObject
    Dim objFullestBag As acObject
    
    'Default settings
    FindPartialBags = False
    
    Set objFullestBag = Nothing
    Set colBags = New colObjects
    
    'Loop through each item in inventory
    For Each objItem In g_Objects.Items.Inv
    
        'if not a salvage bag, skip
        If objItem.itemType <> ITEM_SALVAGE Then GoTo NextItem
        
        If (objItem.UserData(INT_SALVAGECOUNT) > 2) Then GoTo NextItem
        
        'If we already have selected a bag, this one must be of same salvage type
        If Valid(objFullestBag) Then
            If objItem.MaterialType <> objFullestBag.MaterialType Then GoTo NextItem
        'else, if not of desired salvage type, skip
        ElseIf lSalvageType >= 0 Then
            If objItem.MaterialType <> lSalvageType Then GoTo NextItem
        End If
        
        'if bag is complete, skip
        If objItem.UsesLeft >= MAX_UNITS_PER_BAG Then GoTo NextItem
        
        'If bag is not within range of iMinWork to iMaxWork, then skip it
        If (objItem.Workmanship >= iMinWork) And (objItem.Workmanship < (iMaxWork + 1)) Then
            'locDebug "clsSalvager.FindPartialBags: OK OK OK : iMinWork:" & iMinWork & "  iMaxWork:" & iMaxWork
        Else
            'locDebug "clsSalvager.FindPartialBags: SKIPPING : iMinWork:" & iMinWork & "  iMaxWork:" & iMaxWork
            GoTo NextItem
        End If
        
        'If bag not in main inventory or is in exception list...
        If IsImportantItem(objItem) Then GoTo NextItem
        
        'Make sure it's not in the ignore list
        If Valid(colIgnoreList) Then
            If colIgnoreList.Exists(objItem.Guid) Then
                'locDebug "clsSalvager.FindPartialBag - " & objItem.Name & " in ignore list, skipping"
                GoTo NextItem
            End If
        End If
        
        'Make sure this kind of salvage is listed in our salvages list,
        'and that the minimum workmanship is respected
        If Not g_Data.LootFilters.PassFilters(objItem, g_Data.LootFilters.SalvageFilters) Then GoTo NextItem

        locDebug "clsSalvager.FindPartialBag - Adding Bag: " & objItem.Name
        
        'Add bag to collection
        If Not colBags.addObject(objItem) Then
            PrintWarning "clsSalvager.FindPartialBag - Failed to add partial salvage bag " & objItem.Name & " to collection"
            GoTo NextItem
        End If
        
        'If it's the first bag we find...
        If Not Valid(objFullestBag) Then
            Set objFullestBag = objItem
        'If this bag has more salvage units that the previous one...
        ElseIf objItem.UsesLeft > objFullestBag.UsesLeft Then
            Set objFullestBag = objItem
        End If
        
        'Ok we found at least one bag so set function to return true
        FindPartialBags = True
        
NextItem:
    Next objItem

Fin:
    Set objFullestBagOut = objFullestBag
    Set objFullestBag = Nothing
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    FindPartialBags = False
    PrintErrorMessage "clsSalvager.FindPartialBags - " & Err.Description
    Resume Fin
End Function

'Make all the batches of salvages (items & bags) to put on the ust this round
Public Function MakeAllBatches(Optional ByRef colBatchOut As colObjects) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim oFilter As clsLootFilter
    
    Dim colBigBatch As New colObjects 'final list of items & partial bags with different salvage kinds
    Dim colBatch As colObjects        'list of partial bags & items of the same salvage kind to add to the big batch
    Dim oItem As acObject
    Dim iMaxSalWork As Integer
    Dim iMinSalWork As Integer
    
    'Generate the list of valid salvage items currently in inventory
    'Call MakeValidSalvagesList
    
    'Try to make a batch for each salvage kind enabled
    For Each oFilter In g_Data.LootFilters.SalvageFilters
        Set colBatch = Nothing
        
        If oFilter.SalvageCombine Then
            'work at which to lump salvage together
            iMaxSalWork = Val(g_ui.Loot.chSalvageSortWork.Text(g_ui.Loot.chSalvageSortWork.Selected))
        Else
            'Just combine all Work of this salvage type together
            iMaxSalWork = 10
        End If
        
        iMinSalWork = 1

        ' make
        While iMaxSalWork <= 10
            
            If MakeSingleBatch(oFilter.SalvageMaterial, colBatch, iMinSalWork, iMaxSalWork) Then
                'Transfer items in colBatch to colBigBatch
                For Each oItem In colBatch
                    If Not colBigBatch.addObject(oItem) Then
                        PrintErrorMessage "clsSalvage.MakeAllBatches - Failed to add batch item " & oItem.Name & " (id: " & oItem.Guid & ") to Big Batch"
                    End If
                Next oItem
            
                If Not g_ui.Loot.chkEnableMultipleSalvage.Checked Then
                    'If we're not in "multiple salvages at once" mode, exit now
                    bRet = True
                    GoTo Fin
                Else
                    'Found a range of salvage workmaship to lump together, we are done with this salvage type
                    GoTo NextType
                End If
            End If
            iMaxSalWork = iMaxSalWork + 1
            iMinSalWork = iMaxSalWork
        Wend
NextType:
    Next oFilter
    
    'MyDebug "(MakeAllBatch) Content: " & colBigBatch.GetObjectsList
    
    bRet = (colBigBatch.Count > 0)

Fin:
    Set colBatchOut = colBigBatch
    MakeAllBatches = bRet
    Set colBigBatch = Nothing
    Set colBatch = Nothing
    Set oItem = Nothing
    Set oFilter = Nothing
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsSalvager.MakeAllBatches - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function


'Make a batch of salvages (items & bags) to put on the ust
Public Function MakeSingleBatch(ByVal lSalvageMaterial As Long, _
                            Optional ByRef colBatchOut As colObjects, _
                            Optional ByVal salMinWork As Integer = 1, _
                            Optional ByVal salMaxWork As Integer = 10) As Boolean
                            
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim objItem As acObject, objFullestBag As acObject
    
    Dim colBatch As colObjects          'list of bags & items we're going to put in the ust for this round
    Dim colBags As colObjects           'list of bags of the current salvage kind
    Dim colSalv As colObjects           'list of salvageable items of the current salvage kind
    Dim colIgnoreList As colObjects     'list of bags to ignore
    Dim iTotalUnits As Integer, iNumBagsInBatch As Integer
    Dim bFoundBag As Boolean
    
    'Generate the list of valid salvages
    'Call MakeValidSalvagesList

    iTotalUnits = 0
    iNumBagsInBatch = 0
    Set colIgnoreList = New colObjects
    Set colBatch = New colObjects
    Set objFullestBag = Nothing
    
    'First try to complete the current salvage bags if any
    bFoundBag = FindPartialBags(colBags, lSalvageMaterial, colIgnoreList, , salMinWork, salMaxWork)

    If bFoundBag Then
        Dim objBag As acObject
        Dim i As Integer
        Dim SortedBagsList() As acObject    'list of bags sorted by total salvage units descending (highest @ top)

        SortedBagsList = ColToArray(colBags)
        Call SortBySalvageQuantity(SortedBagsList)
            
        Set objFullestBag = SortedBagsList(0)
        iTotalUnits = objFullestBag.UsesLeft
        
        'Start off by adding this bag to the batch since we Complete Biggest Bag First
        If colBatch.addObject(objFullestBag) Then
            ' locDebug "...Biggest Bag: " & objFullestBag.Name & " (" & objFullestBag.UsesLeft & ")"
            iNumBagsInBatch = 1
        Else
            PrintErrorMessage "clsSalvager.MakeBatch - Failed to add objFullestBag '" & objFullestBag.Name & "' to batch collection"
        End If
                    
        'We found a bag, check to see if we can complete it with other bags first
        i = 1
        While ((i <= UBound(SortedBagsList)) And (iTotalUnits < MAX_UNITS_PER_BAG))
            
            'Reference to current bag
            Set objBag = SortedBagsList(i)
            
            'If we can merge the two bags without loosing too much unit in the process...
            If (iTotalUnits + objBag.UsesLeft) <= (MAX_UNITS_PER_BAG + BAGS_MERGE_TOLERANCE) Then
            
                If colBatch.addObject(objBag) Then
                    iTotalUnits = iTotalUnits + objBag.UsesLeft 'increase salvage estimation
                    iNumBagsInBatch = iNumBagsInBatch + 1
                    'locDebug "...Added Bag: " & objBag.Name & " (" & objBag.UsesLeft & ")"
                Else
                    PrintErrorMessage "clsSalvager.MakeBatch - Failed to add partial bag '" & objBag.Name & "' to batch collection"
                End If
                
            End If
    
            'Next bag
            i = i + 1
            
        Wend
            
    End If 'FoundBag

    'If we still have room in the current batch for more salvage
    If iTotalUnits < MAX_UNITS_PER_BAG Then
    
        'Now check if we can salvage some items
        Set colSalv = New colObjects
        
        'For Each objItem In m_colSalvages
        For Each objItem In m_LootedItemsToSalvage
            If (objItem.MaterialType = lSalvageMaterial) _
            And (objItem.Workmanship >= salMinWork) _
            And (objItem.Workmanship <= salMaxWork) Then
                Call colSalv.addObject(objItem)
            End If
        Next objItem
        
        'Found some items of proper salvage type?
        If colSalv.Count > 0 Then
            Dim SortedSalvagesList() As acObject
            
            SortedSalvagesList = ColToArray(colSalv)
            Call SortByWorkmanship(SortedSalvagesList)
            
            'Loop through the sorted salvageable items list
            For i = LBound(SortedSalvagesList) To UBound(SortedSalvagesList)
                
                
                'Is batch full yet?
                If iTotalUnits >= MAX_UNITS_PER_BAG Then
                    GoTo ExitItemsLoop
                End If
                
                'Reference to the current salvage item
                Set objItem = SortedSalvagesList(i)
                
                'Make sure we won't waste too much salvage
                Dim bCanAddToBatch As Boolean
                bCanAddToBatch = (iTotalUnits + objItem.Workmanship <= MAX_UNITS_PER_BAG + BAGS_COMPLETION_TOLERANCE)
                
                'Take into account the special cases for work 8 and above salvage, as there is no way
                'to finish off some bags without losing more than BAGS_COMPLETION_TOLERANCE
                ' For instance, a bag made from 11 work 9 items will be at 99 units and will lose more than
                ' BAGS_COMPLETION_TOLERANCE to complete no matter what you try
                ' ** With Salvaging bonus it's also possible to get 96 units of wk8 salvage, so start there **
                If (iTotalUnits >= 96) And (objItem.Workmanship >= 8) Then
                    bCanAddToBatch = True
                End If
                
                'if for instance we currently only have 1 bag of 99 units of sunstone, and a work 7 Sunstone gem
                'We want to start a new bag with the sunstone item and leave the 99 units bag alone for now
                If (Not bCanAddToBatch) And (colBatch.Count = 1) And (i >= UBound(SortedSalvagesList)) Then
                    Set colBatch = New colObjects
                    iNumBagsInBatch = 0
                    bCanAddToBatch = True
                End If
                
                If (objItem.UserData(INT_SALVAGECOUNT) > 2) Then bCanAddToBatch = False
                
                'Check if we can add it
                If bCanAddToBatch Then
                    If colBatch.addObject(objItem) Then
                        'locDebug "...Added Item: " & objItem.Name & " (" & objItem.Workmanship & ")"
                        iTotalUnits = iTotalUnits + objItem.Workmanship 'increase salvage estimation
                    Else
                        PrintErrorMessage "clsSalvager.MakeBatch - Failed to add salvage item '" & objItem.Name & "' to batch collection"
                        GoTo ExitItemsLoop
                    End If
                End If  'bCanAddToBatch
                
            Next i  'Next item in the sorted salvages list
            
ExitItemsLoop:
        End If  'colSalv > 0
    End If 'Batch complete
    
        
    Dim bValidBatch As Boolean
    If iNumBagsInBatch > 0 Then
        bValidBatch = colBatch.Count >= 2 'at least 2 bags, or 1 bag + item
    Else
        bValidBatch = colBatch.Count >= 1 'can salvage 1 or more items
    End If
    
    'If we made a valid batch, we can leave now
    'If bValidBatch Then locDebug ">>> Valid Batch - Content: " & colBatch.GetObjectsList
    
    Dim tObject As acObject
    Dim tString As String
    For Each tObject In colBatch
        If Valid(tObject) Then
            tString = tString & tObject.Name & " wk(" & tObject.Workmanship & "), "
        End If
    Next tObject
    
    If bValidBatch Then locDebug ">>> Valid Batch: " & tString
    
    'Return if we managed to make a valid batch
    bRet = bValidBatch
    
Fin:
    Set colBatchOut = colBatch
    MakeSingleBatch = bRet
    Set colBatch = Nothing
    Set colBags = Nothing
    Set colSalv = Nothing
    Set colIgnoreList = Nothing
    Set objItem = Nothing
    Set objFullestBag = Nothing
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsSalvager.MakeSingleBatch - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           External Events
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Sub OnIdentifyObject(ByVal objItem As acObject)
On Error GoTo ErrorHandler
    
    If Not Valid(objItem) Then
        PrintErrorMessage "clsSalvager.OnIdentifyObject : invalid objItem - Ignoring."
        GoTo Fin
    End If

    If m_IdQueue.Exists(objItem.Guid) Then
        'Object has been IDed, we can remove it from the ID queue now
        Call m_IdQueue.Remove(objItem.Guid)
        
        'Add object to ItemsToSalvage queue
        locDebug "Object Assessed : " & objItem.Name & " - Description : " & objItem.Description
        If ValidSalvageItem(objItem) Then
            locDebug "OnIdentifyObject - Adding " & objItem.Name & " to Salvage List"
            Call m_ItemsToSalvage.Add(objItem.Guid, objItem.Name)
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.OnIdentifyObject - " & Err.Description
    Resume Fin
End Sub

'You obtain 5 units of Salvaged Leather (workmanship 5.00).
Public Sub OnObtainSalvage()
On Error GoTo ErrorHandler
  
    
'    MyDebug "clsSalvager.OnObtainSalvage - Items salvaged"
'
'    If m_State = ST_WAIT_FOR_COMPLETION Then
'        MyDebug "clsSalvager.OnObtainSalvage - Going back to ST_LOAD_SALVAGES state..."
'        Call SetState(ST_LOAD_SALVAGES)
'    Else
'        PrintWarning "clsSalvager.OnObtainSalvage - Current state is not ST_WAIT_FOR_COMPLETION..."
'    End If
        
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.OnObtainSalvage - " & Err.Description
    Resume Fin
End Sub

'An object has been destroyed, check if it's one of the items we were salvaging
'or an object that was used to recharge a mana stone
Public Sub OnRemoveObject(objItem As acObject)
On Error GoTo ErrorHandler

    If m_State = ST_WAIT_FOR_COMPLETION Then
    
        If Not Valid(m_colBatch) Then
            PrintWarning "clsSalvager.OnRemoveObject - Invalid colBatch"
            GoTo Fin
        End If
        
        'Remove the item from the collection if it was in our batch
        If m_colBatch.Remove(objItem.Guid) Then
            locDebug "clsSalvager.OnRemoveObject - Removing " & objItem.Name & " from batch"
        End If
        
        If m_LootedItemsToSalvage.Remove(objItem.Guid) Then
            locDebug "clsSalvager.OnRemoveObject - Removing " & objItem.Name & " from LootedItemsToSalvage"
        Else
            locDebug "clsSalvager.OnRemoveObject - Could not remove item list Salvage list: " & objItem.Name
        End If
        
        'Are we done usting all the items ?
        If m_colBatch.Count < 1 Then
            If Not g_ui.Options.chkFilterSalvageMsg.Checked Then
                PrintMessage "Salvager: all batch items have been salvaged"
            End If
            Call m_tmrSecure.SetNextTime(3)
            Call SetState(ST_LOAD_SALVAGES)
        End If
        
    ElseIf (objItem.Container = g_Objects.Player.Guid) Then
        
        Call g_Macro.Loot.removeHighManaItem(objItem.Guid)
    
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.OnRemoveObject - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Private Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function OpenUst() As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If UstAvailable Then
        Call g_Service.UseItem(m_objUst)  'UseItemOnSelf
        bRet = True
    End If
    
Fin:
    OpenUst = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSalvager.OpenUst - " & Err.Description
    bRet = False
    Resume Fin
End Function

Private Sub m_tmrTransition_OnTimeout()
    locDebug "clsSalvager.m_tmrTransition_OnTimeout"
    RaiseEvent OnSalvageStopped(m_bComplete)
    Call Reset
End Sub

Public Function ValidSalvageItem(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(objItem) Then
        'salvages must be in the main pack, and not inscribed
        bRet = (objItem.Container = g_Objects.Player.Guid) _
                And (objItem.Description = "") _
                And (Not objItem.IsWielded) _
                And (Not objItem.Equiped) _
                And (Not objItem.HasMajors) _
                And (Not objItem.HasMinors)
    End If
    
    If (objItem.UserData(INT_SALVAGECOUNT) > 2) Then bRet = False
    
Fin:
    ValidSalvageItem = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSalvager.ValidSalvageItem - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Sub ScanSalvages()
On Error GoTo ErrorHandler

    Dim objItem As acObject
    
    'start a new ID queue
    Set m_IdQueue = New colObjects
    
    For Each objItem In g_Objects.Items.Inv
        If (objItem.Container = g_Objects.Player.Guid) And (Not m_ItemsToSalvage.Exists(objItem.Guid)) Then
            Call m_IdQueue.Add(objItem.Guid, objItem.Name)
            'Call g_Service.IDObject(objItem.Guid)
            Call g_Hooks.IDQueueAdd(objItem.Guid)
        End If
    Next objItem
    
    locDebug "ScanSalvages : " & m_IdQueue.GetObjectsList
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.ScanSalvages - " & Err.Description
    Resume Fin
End Sub

Public Sub AddSalvagesToUst()
On Error GoTo ErrorHandler

    Dim objItem As acObject
    
    For Each objItem In m_ItemsToSalvage
        Call g_Hooks.UstAddItem(objItem.Guid)
    Next objItem
    
    Set m_ItemsToSalvage = New colObjects
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSalvager.AddSalvagesToUst - " & Err.Description
    Resume Fin
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub SetState(ByVal iNewState As eSubStates)
    m_State = iNewState
End Sub

Public Function GetStateString() As String
Dim sRes As String

    sRes = ""
    
    Select Case m_State

        Case ST_NONE
            sRes = "None"
            
        Case ST_PREPARE_UST
            sRes = "Preparing Ust"
        
        Case ST_LOAD_SALVAGES
            sRes = "Loading Salvages"
            
        Case ST_DO_SALVAGE
            sRes = "Salvaging"
            
        Case ST_WAIT_FOR_COMPLETION
            sRes = "Waiting for Completion"
            
        Case ST_DONE
            sRes = "Done"
            
        Case Else
            sRes = "[Unknown Salvage State " & m_State & "]"
            
    End Select
    
    GetStateString = sRes
End Function

'Local Debug
Private Sub locDebug(ByVal sDebugMsg As String, Optional ByVal bSilent As Boolean = False)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsSalvager : " & GetStateString & "] " & sDebugMsg, bSilent)
    End If
End Sub


