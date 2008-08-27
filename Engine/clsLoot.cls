VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                            Auto-Looter                                  [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False

'Contants
Private Const LOOT_MAX_SEARCH_CORPSE_TIME = 3
Private Const LOOT_MAX_CORPSE_OPEN_ATTEMPT_TIME = 9
Private Const LOOT_MAX_SEARCH_ITEM_TIME = 9
Private Const LOOT_MAX_PICKUP_TIME = 10 ' seconds max to pickup an item

Private Const DEFAULT_TRANSITION_TIME = 0.25
Private Const MAX_Z_DIFF = 8
Private Const ID_EXPIRATION_TIME = 120      'if the item has been ided in the last 2 minutes,
                                            'dont try to ID it again

'The different looting states
Private Enum eLootStates
    ST_LOOT_SEARCH
    ST_LOOT_OPEN_CORPSE
    ST_LOOT_SEARCH_ITEM
    ST_LOOT_PICKUP_ITEM
    ST_LOOT_ITEM_PICKED_UP
    ST_LOOT_STOP
End Enum

'Members
Private m_State As eLootStates      'current looting state
Private m_curCorpse As acObject     'the corpse we're trying to loot
Private m_curItem As acObject       'the item getting looted
Private m_rareCorpse As acObject    'The corpse that contains a rare

Private m_ItemsIdQueue As colObjects        'list of items left to ID on the body
Private m_ItemsPending As colObjects        'list of items left to receive a OnCreateObject
                                            'NOTE : collection __ONLY__  holding the GUIDs of the items on the corpse
Private m_ItemsOnCorpse As colObjects       'list of the items created on the body
Private m_ItemsToLoot As colObjects         'list of the items we want to loot (items that passed the filters)
Private m_corpseList As colObjects          'list of corpses we need to loot
Private m_colHighManaItems As colObjects    'list of High Mana items we will DESTROY!
Private m_colEmptyManaStones As colObjects   'list of Empty Mana items we will recharge

Private m_bLootCorpses As Boolean       'true if the macro is allowed to loot corpses
Private m_bLootGround As Boolean        'true if macro is allowed to loot items on the ground
Private m_bGroundPickup As Boolean      'true if the item we're picking up was on the ground and not on corpse

Private m_tmrSecure As clsTimer
Private WithEvents m_tmrTransition As clsTimer
Attribute m_tmrTransition.VB_VarHelpID = -1
Private m_tmrNextTime As clsTimer

Private m_dicCorpseToIgnore As Dictionary

Private m_IdQueue As colObjects
Private m_idCount As Integer
Private WithEvents m_tmrExportInv As clsTimer
Attribute m_tmrExportInv.VB_VarHelpID = -1

Private m_NavHold As Long

Private m_silentID As Boolean
Private m_bEmptyManaStone As Boolean

Public Event OnItemLooted(objItem As acObject)
Public Event OnLootingComplete()

Private Sub Reset()

    Set m_ItemsIdQueue = New colObjects

    Set m_ItemsPending = New colObjects
    Set m_ItemsOnCorpse = New colObjects
    Set m_ItemsToLoot = New colObjects

    Set m_curCorpse = Nothing
    Set m_curItem = Nothing
    
    m_bGroundPickup = False
    m_silentID = True
    
    Call m_tmrNextTime.Reset
    Call m_tmrSecure.Reset
    Call m_tmrTransition.Reset
    Call m_tmrExportInv.Reset
    
    Call SetSubState(ST_LOOT_SEARCH)
    
End Sub

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    m_bLootCorpses = True
    m_bLootGround = False
    g_bLootRare = False
    m_bEmptyManaStone = False
    
    Set m_rareCorpse = Nothing
    Set m_tmrSecure = CreateTimer
    Set m_tmrNextTime = CreateTimer
    Set m_tmrTransition = CreateTimer
    Set m_tmrExportInv = CreateTimer
    
    Set m_corpseList = New colObjects
    Set m_IdQueue = New colObjects
    Set m_ItemsIdQueue = New colObjects
    Set m_colHighManaItems = New colObjects
    Set m_colEmptyManaStones = New colObjects
  
    Set m_dicCorpseToIgnore = New Dictionary
    
    m_NavHold = 0

    Call Reset

    m_silentID = False

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.Initialize - Error : " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set m_colHighManaItems = Nothing
    Set m_colEmptyManaStones = Nothing
    
    Set m_curCorpse = Nothing
    Set m_curItem = Nothing
    Set m_IdQueue = Nothing
    Set m_ItemsIdQueue = Nothing
    Set m_ItemsPending = Nothing
    Set m_ItemsOnCorpse = Nothing
    Set m_tmrExportInv = Nothing
    Set m_rareCorpse = Nothing
    Set m_ItemsToLoot = Nothing
    Set m_corpseList = Nothing
    Set m_dicCorpseToIgnore = Nothing
    
    Set m_tmrSecure = Nothing
    Set m_tmrTransition = Nothing
    Set m_tmrNextTime = Nothing
    Set m_tmrExportInv = Nothing
End Sub

'Properties
Public Property Get PickupGUID() As Long
    If Valid(m_curItem) Then
        PickupGUID = m_curItem.Guid
    Else
        PickupGUID = 0
    End If
End Property

Public Property Get PickupItemName() As String
    If Valid(m_curItem) Then
        PickupItemName = m_curItem.Name
    Else
        PickupItemName = "n/a"
    End If
End Property

'------------------------------------------------------------------------------------
' Public Methods
'------------------------------------------------------------------------------------

Public Sub SetSilentID(ByVal aBool As Boolean)
    m_silentID = aBool
End Sub

Public Function silentID() As Boolean
    silentID = m_silentID
End Function

Public Sub idManaStones()
On Error GoTo ErrorMessage
    
    Dim objItem As acObject

    Set m_IdQueue = New colObjects

    For Each objItem In g_Objects.Items.Inv
        If Valid(objItem) Then
            If (objItem.itemType = ITEM_MANA_STONES) And (objItem.Name Like "*Stone*") Then
                'MyDebug "Vitals: id'ing: " & objItem.Name & " (" & objItem.Mana & ")"
                Call m_IdQueue.Add(objItem.Guid, objItem.Name)
                Call g_Hooks.IDQueueAdd(objItem.Guid)
            End If
        End If
    Next objItem

Fin:
    Set objItem = Nothing
    Exit Sub
ErrorMessage:
    PrintErrorMessage "Error in Vitals.idManaStones: " & Err.Description & " - " & Err.Source
    Exit Sub
End Sub

Public Function canLootHighManaItem() As Boolean
    If m_colEmptyManaStones.Count > m_colHighManaItems.Count Then
        canLootHighManaItem = True
    Else
        canLootHighManaItem = False
    End If
End Function

Public Function hasEmptyManaStone() As Boolean
    If m_colEmptyManaStones.Count > 0 Then
        hasEmptyManaStone = True
    Else
        hasEmptyManaStone = False
    End If
End Function

Public Function getEmptyManaStone() As acObject
On Error GoTo ErrorHandler
    Dim objItem As acObject

    If m_colEmptyManaStones.Count <= 0 Then
        Set getEmptyManaStone = Nothing
        GoTo Fin
    End If

Restart:

    Set objItem = Nothing
    For Each objItem In m_colEmptyManaStones
        If Valid(objItem) Then
            If g_Objects.Items.InInventory(objItem.Guid) And (objItem.Mana = 0) And ((g_ds.Time - objItem.LastIdTime) < 30) Then
                Set getEmptyManaStone = objItem
            Else
                Call m_colEmptyManaStones.Remove(objItem.Guid)
                GoTo Restart
            End If
        End If
    Next objItem
    
Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    Set getEmptyManaStone = Nothing
    PrintErrorMessage "clsLoot.getEmptyManaStone - " & Err.Description
    Resume Fin
End Function

Public Sub countEmptyManaStones()
On Error GoTo ErrorMessage
    Dim objItem As acObject

    If m_IdQueue.Count > 0 Then
        'We are still waiting for mana stones to id
        Exit Sub
    End If
    
    Set m_colEmptyManaStones = New colObjects
    
    For Each objItem In g_Objects.Items.Inv
        If Valid(objItem) Then
            If (objItem.itemType = ITEM_MANA_STONES) And (objItem.Name Like "*Stone*") And (objItem.Mana = 0) Then
                MyDebug "countEmptyManaStones: " & objItem.Name & " (" & objItem.Mana & ")"
                Call m_colEmptyManaStones.addObject(objItem)
            End If
        End If
    Next objItem

    MyDebug "countEmptyManaStones: Total : " & m_colEmptyManaStones.Count
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorMessage:
    PrintErrorMessage "Error in clsLoot.countEmptyManaStones: " & Err.Description & " - " & Err.Source
    Exit Sub
End Sub


Public Function hasHighManaItem() As Boolean
    If m_colHighManaItems.Count > 0 Then
        hasHighManaItem = True
    Else
        hasHighManaItem = False
    End If
End Function

Public Function getHighManaItem() As acObject
On Error GoTo ErrorHandler
    Dim objItem As acObject

    If m_colHighManaItems.Count <= 0 Then
        Set getHighManaItem = Nothing
        GoTo Fin
    End If

Restart:

    Set objItem = Nothing
    For Each objItem In m_colHighManaItems
        If Valid(objItem) Then
            If g_Objects.Items.InInventory(objItem.Guid) Then
                'g_Objects.FindObject(objItem.GUID,False)
                Set getHighManaItem = objItem
            Else
                Call m_colHighManaItems.Remove(objItem.Guid)
                GoTo Restart
            End If
        End If
    Next objItem
    
Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    Set getHighManaItem = Nothing
    PrintErrorMessage "clsLoot.getHighManaItem - " & Err.Description
    Resume Fin
End Function

Public Sub removeHighManaItem(ByVal aGuid As Long)
On Error GoTo ErrorHandler

    If m_colHighManaItems.Count <= 0 Then
        Exit Sub
    End If

    If Valid(g_Macro.HighManaTarget) Then
        If g_Macro.HighManaTarget.Guid = aGuid Then
            g_Macro.HighManaTarget = Nothing
        End If
    End If

    If m_colHighManaItems.Exists(aGuid) Then
        Call m_colHighManaItems.Remove(aGuid)
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.removeHighManaItem - " & Err.Description
    Resume Fin
End Sub

Public Function canLoot() As Boolean
On Error GoTo ErrorHandler

    If m_corpseList.Count > 0 Then
        canLoot = True
    Else
        canLoot = False
    End If
    
Fin:
    Exit Function
ErrorHandler:
    canLoot = False
    PrintErrorMessage "clsLoot.canLoot - " & Err.Description
    Resume Fin
End Function

Public Sub incNavHold()
    'PrintErrorMessage ("inc'd")
    m_NavHold = m_NavHold + 1
End Sub

Public Sub decNavHold()
    'PrintErrorMessage ("dec'd")
    If m_NavHold > 0 Then
       m_NavHold = m_NavHold - 1
    Else
       m_NavHold = 0
    End If
End Sub

Public Function getNavHold()
    getNavHold = m_NavHold
End Function

Public Sub resetNavHold()
    m_NavHold = 0
End Sub

Public Function StartLooting(Optional bLootCorpse As Boolean = True, Optional bLootGround As Boolean = False) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean

    m_bLootCorpses = bLootCorpse
    m_bLootGround = bLootGround
    
    If Valid(m_rareCorpse) Then
        'We have a rare corpse to loot!
        Call Reset
        bRet = True
        Set m_curCorpse = m_rareCorpse
        Call SetSubState(ST_LOOT_OPEN_CORPSE)
        MyDebug "clsLoot.StartLooting: found a high priority corpse with a rare"
        GoTo Fin
    End If
    
    'check if we can loot and if there are corpses to loot
    If (Not g_Objects.Items.BackpackFull) _
    And (FoundCorpseToLoot(m_curCorpse) Or FoundGroundItemToLoot(m_curItem)) Then
        Call Reset
        bRet = True
    Else
        bRet = False
    End If
    
Fin:
    StartLooting = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLoot.StartLooting - " & Err.Description
    Resume Fin
End Function

Private Sub m_tmrTransition_OnTimeout()
    Call Reset
    RaiseEvent OnLootingComplete
End Sub

Public Sub StopLooting(Optional sSrc As String = "", Optional dTransitionTime As Double = DEFAULT_TRANSITION_TIME)
On Error GoTo ErrorHandler
    
    locDebug "clsLoot.StopLooting - From: " & sSrc
    Call m_tmrTransition.SetNextTime(dTransitionTime)
    g_bLootRare = False
    Set m_rareCorpse = Nothing
    
    'Check to see if we have a bad egg that needs to be removed
    If Valid(m_curCorpse) Then
        If (m_curCorpse.UserData(INT_BLISTCOUNT) > 2) Then
            Call m_curCorpse.SetUserData(B_LOOTED, True)
            MyDebug "clsLoot.StopLooting: INT_BLISTCOUNT > 2 for " & m_curCorpse.Name
            MyDebug "clsLoot.StopLooting: marking as looted to remove from list"
        Else
            Dim iCount As Integer
            iCount = m_curCorpse.UserData(INT_BLISTCOUNT) + 1
            Call m_curCorpse.SetUserData(INT_BLISTCOUNT, iCount)
        End If
    End If
    
    Call cleanCorpseList
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.StopLooting - " & Err.Description
    Resume Fin
End Sub

'Clear out m_corpseList
Private Sub cleanCorpseList()
On Error GoTo Error_Handler
    
    Dim tmpObj As acObject
    Dim newList As colObjects
    
    Set newList = New colObjects
    
    For Each tmpObj In m_corpseList
        If (tmpObj.itemType = ITEM_CORPSE) Then
            If tmpObj.UserData(B_LOOTED) Then
                'do nothing (will be removed)
            ElseIf (g_Core.Time > tmpObj.UserData(L_TIME)) Then
                'do nothing (will be removed)
            ElseIf (tmpObj.GetSquareRange <= (g_Data.LootSearchRadius * g_Data.LootSearchRadius)) Then
                Call newList.addObject(tmpObj)
            End If
        End If
    Next tmpObj
    
    Set m_corpseList = newList

Fin:
    Set tmpObj = Nothing
    Set newList = Nothing
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.cleanCorpseList - " & Err.Description
    Resume Fin
End Sub

'Called by OnSetPackContent
Public Sub OnOpenContainer(ByVal objContainer As acObject, ByVal colItems As colObjects)
On Error GoTo Error_Handler
    
    If Not Valid(m_curCorpse) Then
        PrintWarning "WARNING - OnOpenContainer : invalid m_curCorpse"
        Exit Sub
    End If
    
    locDebug "clsLoot.OnOpenContainer - " & colItems.GetObjectsList
        
    'make sure it's the body we're trying to loot
    If m_curCorpse.Guid = objContainer.Guid Then
    
        'save the items pending list
        Set m_ItemsPending = colItems
        
        'the corpse is now opened
        Call OnCorpseOpened
        
    Else
        locDebug "clsLoot.OnOpenContainer - The corpse opened (" & objContainer.Name & ") is not the corpse being looted by the macro (" & m_curCorpse.Name & ") !"
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.OnOpenContainer - " & Err.Description
    Resume Fin
End Sub

Public Function HandleActionFailure(ByVal iFailureId As Integer) As Boolean
On Error GoTo Error_Handler
    
    Dim bRet As Boolean
    
    locDebug "HandleActionFailure - iFailureID: " & iFailureId
    
    Select Case iFailureId
        Case FAIL_UNABLE_TO_MOVE_TO_OBJECT, _
                FAIL_ALREADY_IN_USE, _
                FAIL_CANNOT_BE_USED, _
                FAIL_CANT_OPEN_BODY
            
            If Valid(m_curCorpse) Then
                locDebug "HandleActionFailure - Adding " & m_curCorpse.Name & " to corpse ignore list"
                Call AddCorpseToIgnore(m_curCorpse)
            End If
            
            Call StopLooting("HandleActionFailure - Unable to reach body/pick up item - " & iFailureId)
            bRet = True
            
        Case Else
            bRet = False
    End Select
    
Fin:
    HandleActionFailure = bRet
    Exit Function
Error_Handler:
    PrintErrorMessage "clsLoot.HandleActionFailure - " & Err.Description
    Resume Fin
End Function

Private Sub AddItemToLoot(objItem As acObject)
On Error GoTo ErrorHandler

    If Valid(objItem) Then
        If Not m_ItemsToLoot.Exists(objItem.Guid) Then
            locDebug "AddItemToLoot - Added: " & objItem.Name
            Call m_ItemsToLoot.addObject(objItem)
        Else
            PrintWarning "clsLoot.AddItemToLoot - " & objItem.Name & " is already in m_ItemsToLoot. Ignoring."
        End If
    Else
        PrintErrorMessage "clsLoot.AddItemToLoot - invalid objItem"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.AddItemToLoot - " & Err.Description
    Resume Fin
End Sub

'ValidateItem - Apply the pickup filters to check if this item can
'make it to the m_ItemsToLoot list or not
Private Function ValidateItem(ByVal objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(objItem) Then
        If ValidPickup(objItem) Then
            
            'Add this item to our items-to-loot list
            Call AddItemToLoot(objItem)
            
            'If we're currently looking for items to loot on the corpse,
            'tell the macro we can perform a new search right now
            If m_State = ST_LOOT_SEARCH_ITEM Then
                Call m_tmrNextTime.ExpireNow
            End If
            
            bRet = True
        End If
    Else
        PrintErrorMessage "clsLoot.ValidateItem - invalid objItem"
    End If
    
Fin:
    ValidateItem = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.ValidateItem - " & Err.Description
    bRet = False
    Resume Fin
End Function

Private Function AddItemToCorpse(objItem As acObject) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    bRet = False
    
    If Not Valid(objItem) Then
        PrintErrorMessage "clsLoot.AddItemToCorpse - invalid objItem - Ignoring"
        GoTo Fin
    End If
    
    'ID the item if we have some filters enabled requiring it
    'SPK - Fixed bad branching where items not in the CanAssess
    '      list wouldn't be able to pass the filters
    If MustIdItems And CanAssess(objItem) Then
        locDebug "Added " & objItem.Name & " to Loot Id Queue..."
        
        'Extend security timer for ID time
        Call m_tmrSecure.AddTime(2)
        
        'add to pending IDs collection
        Call m_ItemsIdQueue.Add(objItem.Guid, objItem.Name)
        
        'ask for an ID on this item
        'Call g_Service.IDObject(objItem.Guid)
        Call g_Hooks.IDQueueAdd(objItem.Guid)
        
    Else
        locDebug "Object moved to corpse : " & objItem.Name
        
        'Directly check if the item can be picked up
        Call ValidateItem(objItem)
    End If
    
    'If this item was in the "object creation pending" list (ItemsPending collection)
    'remove it from there since it has now been created
    Call m_ItemsPending.Remove(objItem.Guid)
    
    'Move this item to the ItemsOnCorpse collection
    Call m_ItemsOnCorpse.addObject(objItem)
    
    'If we're currently looking for items to loot on the corpse,
    'tell the macro we can perform a new search right now
    If m_State = ST_LOOT_SEARCH_ITEM Then
        Call m_tmrNextTime.ExpireNow
    End If

    'return ok
    bRet = True

Fin:
    AddItemToCorpse = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.AddItemToCorpse - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Sub OnCreateObject(ByVal objItem As acObject)
On Error GoTo Error_Handler
  
    If Not Valid(objItem) Then
        PrintErrorMessage "clsLoot.OnCreateObject : invalid objItem - Ignoring."
        Exit Sub
    End If
    
    locDebug "clsLoot.OnCreateObject: " & objItem.Name & "  type: " & objItem.itemType & " dataType: " & objItem.gameDataType
    
    'if it's a corpse, ask for ID but don't add to ID queue
    'this corpse ID is MANDATORY, as it checks if the macro has rights on this body
    If objItem.itemType = ITEM_CORPSE Then
        locDebug "OnCreateObject : " & objItem.Name & " created - Requesting ID..."
        
        'Call g_Service.IDObject(objItem.Guid)
        Call g_Hooks.IDQueueAdd(objItem.Guid)
        
        ' Set a timer on the corpse, so we don't loot corpses over 5 mins old
        Call objItem.SetUserData(L_TIME, g_Core.Time + CORPSE_TIMER)
        
        'Add corpse to List
        Call m_corpseList.addObject(objItem)
        
        locDebug "clsLoot.OnCreateObject: " & objItem.Name & " (" & m_corpseList.Count & ")"
        
        GoTo Fin
    
    'If we're in looting mode and this object belongs to the
    'object list we received from OnOpenCorpse
    ElseIf (g_Macro.State = ST_LOOT) And m_ItemsPending.Exists(objItem.Guid) Then
        locDebug "clsLoot.OnCreateObject: ST_LOOT, adding item to corpse"
        Call AddItemToCorpse(objItem)
    End If
    
Fin:
    locDebug "clsLoot.OnCreateObject Fin"
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.OnCreateObject - " & Err.Description
    Resume Fin
End Sub

'Called when an object has just been identified/assessed
Public Sub OnIdentifyObject(ByVal objItem As acObject)
On Error GoTo Error_Handler
        
    If Not Valid(objItem) Then
        PrintErrorMessage "clsLoot.OnIdentifyObject : invalid objItem - Ignoring."
        GoTo Fin
    End If
    
    'Send rare stats.
    If g_ui.Loot.chkSendRares.Checked And objItem.IsRare Then
        MyDebug "Getting ready to send rare stats..."
        Dim lowDmg As Long
        lowDmg = objItem.GetLowDamage
        
        Call g_RareTracker.SendStats(objItem.Name, objItem.RareNumber, objItem.ActivateSkill, objItem.ActivateSkillVal, objItem.ArmorLevel, objItem.ArmorType, objItem.AssociatedSpellId, objItem.AttackBonus, objItem.Burden, objItem.Coverage3, objItem.DamageFlags, objItem.DamageModifier, objItem.DefenseBonus, objItem.Description, objItem.ElementBonusDamage, lowDmg, objItem.HighDamage, objItem.Imbue, objItem.LoreReq, objItem.MagicDefense, objItem.MaxMana, objItem.ManaConvMod, objItem.MaterialType, objItem.MissileDefense, objItem.PvMBonus, objItem.RaceReq, objItem.RankReq, objItem.SkillReqID, objItem.SkillUsed, objItem.Spells, objItem.Spellcraft, objItem.TotalUses, objItem.UsesLeft, objItem.Value, objItem.Variance, objItem.WieldReqID, objItem.WieldReqType, objItem.WieldReqVal, objItem.Workmanship, objItem.BitingStrike, objItem.CrushingBlow, objItem.ShortDesc, objItem.UsageInstructions)
    End If
        
    If Not g_ui.Options.chkFilterIdMsg.Checked And Not m_silentID And _
        ((g_Hooks.CurrentSelection = objItem.Guid) Or (objItem.Wielder <> 0)) Then
            locDebug "ID'ing selected item"
            Call printIdString(objItem)
            If g_Macro.Running Then
                Call SetSilentID(True)
            End If
    End If
    
    If m_IdQueue.Count > 0 Then
        If m_IdQueue.Exists(objItem.Guid) Then
            Call m_IdQueue.Remove(objItem.Guid)
            Exit Sub
        End If
    End If

    If (objItem.itemType = ITEM_CORPSE) And (objItem.KillerName = g_ds.Player.Name) Then
        If InStr(LCase(objItem.Description), "rare") Then
            g_bLootRare = True
            Set m_rareCorpse = objItem
            Call m_rareCorpse.SetUserData(B_HASRARE, True)
            MyDebug "clsLoot.OnIdentifyObject: found corpse containing a rare"
        End If
        
        'Call incNavHold
    End If
    
    If m_ItemsIdQueue.Count > 0 Then
        If m_ItemsIdQueue.Exists(objItem.Guid) Then
            locDebug "Object assessed : " & objItem.Name & " - Value:" & objItem.Value & " - Burden: " & objItem.Burden & " - Type: " & objItem.itemType
    
            'Object has been IDed, we can remove it from the ID queue now
            Call m_ItemsIdQueue.Remove(objItem.Guid)
        
            'Check if item is lootable
            Call ValidateItem(objItem)
        
            'If we're currently looking for items to loot on the corpse,
            'tell the macro we can perform a new search right now
            If m_State = ST_LOOT_SEARCH_ITEM Then Call m_tmrNextTime.ExpireNow
        End If
    End If
    
    If (objItem.itemType = ITEM_MANA_STONES) And (objItem.Name Like "*Stone*") Then
        Call g_Macro.Loot.countEmptyManaStones
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.OnIdentifyObject - " & Err.Description
    Resume Fin
End Sub

Public Sub OnRemoveObject(ByVal objItem As acObject)
On Error GoTo Error_Handler

    locDebug "clsLoot.OnRemoveObject: " & objItem.Name & " (" & m_corpseList.Count & ")"
    'Call m_corpseList.Remove(objItem.Guid)
    Call objItem.SetUserData(B_LOOTED, True)

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.OnRemoveObject - " & Err.Description
    Resume Fin
End Sub

' We get a Rare found message
Public Sub OnRareFound(ByVal aMessage As String)
On Error GoTo Error_Handler
    
    ' aMessage will be of the form:
    '         Xeon Xarid has discovered the Dodger's Crystal
    ' so we need to parse out the name of the item discovered, so we can try and loot it
    
    Dim aMatch As String
    Dim aName As String
    Dim iPos As Integer
    Dim aRegex As New RegExp
    Dim colMatches As MatchCollection
    Dim objRareFinder As acObject
    Dim NewPlayerGUID As String
            
    aRegex.Pattern = "has discovered the (.+)"
    aRegex.Global = True
    aRegex.IgnoreCase = True
    
    If aRegex.Test(aMessage) Then
        Set colMatches = aRegex.Execute(aMessage)
        aMatch = colMatches.Item(0).SubMatches(0)
        'Remove ! from the end
        If InStrRev(aMatch, "!") Then
            iPos = InStrRev(aMatch, "!")
            aName = Left(aMatch, iPos - 1)
        Else
            aName = aMatch
        End If
        
        MyDebug "OnRareFound: rare name: " & aName
        
        MyDebug "Checking to see who found the rare..."
        
        If InStr(aMessage, g_ds.Player.Name) = 0 Then 'Someone else found the rare.
            Dim HadDiscTextStartPos As Integer
            Dim playerName As String
            
            HadDiscTextStartPos = InStr(aMessage, "has discovered the") ' Position of has discovered text.
                        
            playerName = Left(aMessage, HadDiscTextStartPos - 1) 'Get the player name.
            
            MyDebug "Someone else (" & playerName & ") found the rare."
            
            Set objRareFinder = g_Objects.FindPlayerByName(playerName)
            
            NewPlayerGUID = objRareFinder.Guid
            
            MyDebug "Checking to see if RareTracker is enabled..."
            
            If g_ui.Loot.chkSendRares.Checked Then
                MyDebug "RareTracker Enabled. Sending rare information."
                Call g_RareTracker.SendData(aName, playerName, NewPlayerGUID)
            End If
        Else 'We found the rare.
            MyDebug "We found the rare."
            MyDebug "Checking to see if RareTracker is enabled..."
            
            If g_ui.Loot.chkSendRares.Checked Then
                MyDebug "RareTracker Enabled. Sending rare information."
                Call g_RareTracker.SendData(aName, g_ds.Player.Name, "0")
            End If
        End If

        'If Not g_ui.Loot.ItemInPickupList(aName) Then
        '    Call g_ui.Loot.AddItemToList(g_ui.Loot.lstPickup, aName, True, True)
        '    PrintMessage "Adding rare to pickup list: " & aName
        'End If
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.OnRareFound - " & Err.Description & " " & Erl
    Resume Fin
End Sub

Public Function CheckPickup(ObjGUID As Long, Optional bOnAdjustStackSize As Boolean = False) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim objItem As acObject
    
    'default
    bRet = False
    
    If Not Valid(m_curItem) Then
        GoTo Fin
    End If
        
    If g_Objects.Items.Exists(ObjGUID, objItem) Then
        If (objItem.Guid = PickupGUID) Or (bOnAdjustStackSize And SameText(objItem.Name, PickupItemName)) Then
            Call objItem.SetUserData(B_MACO_PICKUP, True)
            Call OnItemPickup(m_curItem)
            bRet = True
        End If
    End If

Fin:
    CheckPickup = bRet
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLoot.CheckPickup - " & Err.Description
    Resume Fin
End Function

'------------------------------------------
' Macro Loot State - Loot Engine
'------------------------------------------
Public Sub RunState()
On Error GoTo Error_Handler
   
1    If m_tmrTransition.Enabled Then GoTo Fin
    
2    If m_tmrSecure.Expired Then
        PrintWarning "clsLoot.RunState[" & GetStateString & "] - Secure Timer expired : exiting state."
        Call StopLooting("RunState - Secure Timer Expired")
        GoTo Fin
    End If
   
    'Staying alive is more important than looting.
    'However, still keep this body on the list of corpses to loot
3    If g_Macro.NeedHealing("clsLoot.RunState") Then
        MyDebug "RunState - Healing..."
        Call StopLooting("RunState - Need Healing")
        GoTo Fin
    End If
    
    'Look for targets
    If Not g_ui.Loot.chkBoostLootPriority.Checked _
    And Not g_ui.Loot.chkEnableLooting.Checked _
    And Not g_bLootRare _
    And (g_Macro.Combat.Target Is Nothing) Then
        If FindBestTarget() Then
            locDebug "clsLoot.Runstate : found target, moving to regular IDLE state for combat initialization"
            'go back to regular IDLE mode to ensure we're in the proper equipment/vital state
            Call SetSubState(ST_LOOT_STOP)
            Call StopLooting("RunState - Found Combat Target")
        End If
    End If

5    If m_tmrNextTime.Expired Then
        
'        If g_ui.Loot.chkLootForcePeaceMode.Checked And (g_Hooks.CombatMode <> COMBATSTATE_PEACE) Then
'            Call g_Macro.RequestCombatStateChange(COMBATSTATE_PEACE)
'            Exit Sub
'        End If
        
        Select Case m_State
        
            Case ST_LOOT_SEARCH
6                If FoundGroundItemToLoot(m_curItem) Then

                    locDebug "Found GROUND item " & m_curItem.Name & ". Now trying to pick it up."
                    m_bGroundPickup = True
                    If Not g_ui.Options.chkFilterLootMsg.Checked Then
                        PrintMessage "Looting ground item : " & m_curItem.Name
                    End If
                    Call SetSubState(ST_LOOT_PICKUP_ITEM)
                
7                ElseIf FoundCorpseToLoot(m_curCorpse) Then
                    locDebug "Found Corpse : " & m_curCorpse.Name & ". Now trying to open it..."
                    
                    'start a clean items ID queue list for the new body
                    Set m_ItemsIdQueue = New colObjects
                    
                    'try to open the body
                    Call SetSubState(ST_LOOT_OPEN_CORPSE)
                Else
                    'Stop looting if nothing to loot
                    Call StopLooting("RunState - ST_LOOT_SEARCH")
                End If
                
            Case ST_LOOT_OPEN_CORPSE
8               If Not DoOpenCorpse Then
                    locDebug "Couldn't open corpse (unreachable or invalid). Searching for Corpse again."
                    Call SetSubState(ST_LOOT_SEARCH)
                Else
                    locDebug "Corpse selected - waiting for Corpse to open."
                    'Should then receive an OnCorpseOpen event, enabling ST_LOOT_SEARCH_ITEM state
                End If
            
            'Search for items to loot on the body
            'Or stop looting if there aren't any items left
            Case ST_LOOT_SEARCH_ITEM
                
9                If Not Valid(m_curCorpse) Then
                   locDebug "WARNING - ST_LOOT_SEARCH_ITEM - invalid m_curCorpse - Searching for Corpse again."
                   Call SetSubState(ST_LOOT_SEARCH)
                
                ElseIf (m_curCorpse.GetSquareRange > (g_Data.LootSearchRadius * g_Data.LootSearchRadius)) Then
                   locDebug m_curCorpse.Name & " is out of range, searching for another one"
                   Call SetSubState(ST_LOOT_SEARCH)
                   
                ElseIf g_Objects.Items.BackpackFull Then    'make sure we can still pickup items
                    PrintMessage "Backpack is full - can't pickup anymore items."
                    Call StopLooting("RunState - ST_LOOT_SEARCH_ITEM(Backpackfull)")
                    
                ElseIf FoundItemToPickup(m_curItem) Then
                    Dim sLootString As String
        
                    sLootString = "Looting " & m_curItem.Name
                    
                    If m_curItem.Workmanship <> 0 Then sLootString = sLootString & " [Work: " & m_curItem.Workmanship & "]"
                    If g_ui.Loot.chkPickupValuable.Checked Then sLootString = sLootString & " [Value: " & m_curItem.Value & "]"
                    
                    If g_ui.Loot.chkBurdenRatio.Checked And (m_curItem.Burden <> -1) And (m_curItem.Burden <> 0) Then
                        sLootString = sLootString & " [Burden: " & m_curItem.Burden & " - Ratio: " & m_curItem.Value / m_curItem.Burden & " ]"
                    End If
                    
                    ' JSC -- filter check
                    If Not g_ui.Options.chkFilterLootMsg.Checked Then
                        PrintMessage sLootString
                    Else
                        locDebug "clsLoot.RunState: " & sLootString
                    End If
                    Call SetSubState(ST_LOOT_PICKUP_ITEM)
                Else
                    'make sure we don't have item IDs pending on this corpse before closing it
                    If m_ItemsIdQueue.Count > 0 Then
                        locDebug "Waiting for pending loot IDs on : " & m_ItemsIdQueue.GetObjectsList
                        Call m_tmrNextTime.SetNextTime(2)
                    'make sure we don't have any items waiting for creation on this corpse
                    ElseIf m_ItemsPending.Count > 0 Then
                        Dim strPending As String
                        Dim objItem As acObject
                    
                        strPending = ""
                        For Each objItem In m_ItemsPending
                            'if we somehow missed the OnCreateObject event for this object...
                            If g_Objects.Items.Exists(objItem.Guid) Then
                                locDebug "ST_LOOT_SEARCH_ITEM - Found a pending item already created : " & objItem.Name
                                Call AddItemToCorpse(objItem)
                                GoTo Fin
                            Else
                                strPending = strPending & objItem.Name & ", "
                            End If
                        Next objItem
                    
                        locDebug "Waiting for pending corpse objects creation : " & strPending
                        'locDebug "Waiting for pending corpse objects creation : " & m_ItemsPending.GetObjectsList
                        Call m_tmrNextTime.SetNextTime(2)
                    Else 'nothing more to loot, close body
                        MyDebug "No more items to pickup on " & m_curCorpse.Name & ". Stopping looting."
                        Call SetBodyLooted
                        Call SetSubState(ST_LOOT_STOP)
                        Call StopLooting("RunState - ST_LOOT_SEARCH_ITEM")
                        If PhatLoot.doAutoStacking Then
                            locDebug "Called PhatLoot.doAutoStacking with return of TRUE"
                        End If
                    End If
                End If
            
            'Attempt to pickup the item
            Case ST_LOOT_PICKUP_ITEM
     
                'Try to pickup the item
                If g_Hooks.BusyState <> 4 Then
                    
'                    If g_ui.Loot.chkLootForcePeaceMode.Checked And (g_Hooks.CombatMode <> COMBATSTATE_PEACE) Then
'                        Call g_Macro.RequestCombatStateChange(COMBATSTATE_PEACE)
'                        Exit Sub
'                    End If
                
                    If Not DoPickupItem Then
                        'couldn't pickup item, move back to search item
                        If m_bGroundPickup Then
                            Call StopLooting("RunState[ST_LOOT_PICKUP_ITEM] - couldn't reach ground item")
                        Else
                            locDebug "Unable to pickup the current item. Searching for another one"
                            Call SetSubState(ST_LOOT_SEARCH_ITEM)
                        End If
                    End If
                    
                End If
            
            'The item was successfully picked up
            Case ST_LOOT_ITEM_PICKED_UP
                If Valid(m_curItem) Then
13                locDebug m_curItem.Name & " successfully picked up. Searching for another item to pickup."
                End If
    
                '1 second delay before next pickup
                Call m_tmrNextTime.SetNextTime(1)
                
                If m_bGroundPickup Then
                    Call StopLooting("ST_LOOT_ITEM_PICKED_UP")
                Else
                    'Go back to SearchItem state
                    Call SetSubState(ST_LOOT_SEARCH_ITEM)
                End If
            
            Case Else
                locDebug "RunState: unknown state " & m_State
                
        End Select
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.RunState - " & Err.Description & " - line: " & Erl
    Resume Fin
    
End Sub

'------------------------------------------------------------------------------------
' Private Methods
'------------------------------------------------------------------------------------

'Changes the current Looting State
Private Function SetSubState(ByVal iNewState As Integer)
    'MyDebug "SetLootState: " & iNewState, True
    m_State = iNewState
    
    'Add error checking timer for each state
    Select Case m_State
        Case ST_LOOT_SEARCH
            Call m_tmrSecure.SetNextTime(LOOT_MAX_SEARCH_CORPSE_TIME)
    
        Case ST_LOOT_OPEN_CORPSE
            Call m_tmrSecure.SetNextTime(LOOT_MAX_CORPSE_OPEN_ATTEMPT_TIME)
            
        Case ST_LOOT_SEARCH_ITEM
            Call m_tmrSecure.SetNextTime(LOOT_MAX_SEARCH_ITEM_TIME + 2 * m_ItemsIdQueue.Count)
            
        Case ST_LOOT_PICKUP_ITEM
            Call m_tmrSecure.SetNextTime(LOOT_MAX_PICKUP_TIME)
        
        Case Else
            Call m_tmrSecure.SetNextTime(8)
    End Select
    
End Function

Public Function IsValidCorpse(objCorpse As acObject) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    If Valid(objCorpse) Then
        If (objCorpse.itemType = ITEM_CORPSE) Then
            bRet = ((objCorpse.IsOnGround) And CanLootCorpse(objCorpse))
        Else
            bRet = False
        End If

        If (objCorpse.itemType = ITEM_CORPSE) Then
            locDebug "IsValidCorpse: IsOnGround: " & objCorpse.IsOnGround & "  : " & objCorpse.Name & " : " & objCorpse.Guid
            locDebug "IsValidCorpse: UserData(B_LOOTED): " & objCorpse.UserData(B_LOOTED)
            locDebug "IsValidCorpse: UserData(L_TIME): " & objCorpse.UserData(L_TIME) & ":" & g_Core.Time
        End If
        
        If g_ui.Loot.chkFilterCorpses.Checked And bRet Then
            bRet = g_ui.Loot.ItemInCorpseList(objCorpse.Name, True)
        End If
    Else
        bRet = False
    End If
    
Fin:
    IsValidCorpse = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLoot.IsValidCorpse - " & Err.Description
    Resume Fin
End Function

Private Function CanLootCorpse(objCorpse As acObject) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    bRet = False
    
    If Not Valid(objCorpse) Then
        PrintErrorMessage "clsLoot.CanLootCorpse - invalid objCorpse - Ignoring"
        Exit Function
    End If
    
    If Not (objCorpse.itemType = ITEM_CORPSE) Then
        locDebug "CanLootCorpse - " & objCorpse.Name & " is NOT a corpse"
        bRet = False
    ElseIf MustIgnoreCorpse(objCorpse) Then
        locDebug "CanLootCorpse - " & objCorpse.Name & " is on CorpseToIgnore list"
        bRet = False
    ElseIf objCorpse.UserData(B_LOOTED) Then
        locDebug "CanLootCorpse - " & objCorpse.Name & " has already been marked as looted"
        bRet = False
    ElseIf (g_Core.Time > objCorpse.UserData(L_TIME)) Then
        locDebug "CanLootCorpse - " & objCorpse.Name & " is to old: curTime:" & g_Core.Time & ":L_TIME:" & objCorpse.UserData(L_TIME)
        bRet = False
    ElseIf SameText(objCorpse.KillerName, g_Objects.Player.Name) Then
        bRet = True
    ElseIf g_ui.Loot.chkLootFellowshipKills.Checked And g_Objects.Fellowship.Active Then
        If g_Objects.Fellowship.PlayerShareLoot(objCorpse.KillerName) Then
            locDebug "CanLootCorpse - Found valid corpse to loot, killed by fellow " & objCorpse.KillerName
            bRet = True
        End If
    End If
    
Fin:
    CanLootCorpse = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsLoot.CanLootCorpse - " & Err.Description
    Resume Fin
End Function

'Looks up for items on ground
Private Function FoundGroundItemToLoot(Optional ByRef objItemOut As acObject) As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    Dim tmpObj As acObject

    'default return values
    bRet = False
    Set objItemOut = Nothing
    
    If Not m_bLootGround Then
        GoTo Fin
    End If
    
    For Each tmpObj In g_Objects.Items.World
        If ValidPickup(tmpObj) And (tmpObj.IsOnGround) Then
            If (tmpObj.GetSquareRange <= (g_Data.LootSearchRadius * g_Data.LootSearchRadius)) _
            And (ZDiff(tmpObj, g_Objects.Player) <= MAX_Z_DIFF) Then
                locDebug "FoundGroundItemsToLoot : found " & tmpObj.Name
                Set objItemOut = tmpObj
                bRet = True
                GoTo Fin
            Else
                locDebug "FoundGroundItemToLoot: Fail - found " & tmpObj.Name & " but TOO FAR."
            End If
        'Else
        '    locDebug "FoundGroundItemToLoot: Unlootable Corpse : " & tmpObj.Name
        End If
    Next tmpObj

Fin:
    FoundGroundItemToLoot = bRet
    Set tmpObj = Nothing
    Exit Function
Error_Handler:
    PrintErrorMessage "clsLoot.FoundGroundItemsToLoot - " & Err.Description
    Set objItemOut = Nothing
    bRet = False
    Resume Fin
End Function

'Looks up for bodies to loot around us. Returns true if it has found one
Private Function FoundCorpseToLoot(Optional ByRef objCorpseOut As acObject) As Boolean
On Error GoTo Error_Handler
    Dim bRet As Boolean
    Dim tmpObj As acObject

    'default return values
    bRet = False
    Set objCorpseOut = Nothing
    
    If Not m_bLootCorpses Then
        Exit Function
    End If
    
    For Each tmpObj In m_corpseList
        If IsValidCorpse(tmpObj) Then
            If (tmpObj.GetSquareRange <= (g_Data.LootSearchRadius * g_Data.LootSearchRadius)) _
            And (ZDiff(tmpObj, g_Objects.Player) <= MAX_Z_DIFF) Then
                locDebug "FoundCorpseToLoot : found " & tmpObj.Name & " : " & tmpObj.Guid
                Set objCorpseOut = tmpObj
                bRet = True
                GoTo Fin
            Else
                locDebug "FoundCorpseToLoot: " & tmpObj.Name & " is TOO FAR away"
                'Call m_corpseList.Remove(tmpObj.Guid)
            End If
        'Else
        '    locDebug "FoundCorpseToLoot: Unlootable Corpse : " & tmpObj.Name
        End If
    Next tmpObj

Fin:
    FoundCorpseToLoot = bRet
    Set tmpObj = Nothing
    Exit Function
Error_Handler:
    PrintErrorMessage "clsLoot.FoundCorpseToLoot - " & Err.Description
    bRet = False
    Set objCorpseOut = Nothing
    Resume Fin
End Function

Private Sub OnCorpseOpened()
On Error GoTo ErrorHandler

    If Not Valid(m_curCorpse) Then
        PrintErrorMessage "clsLoot.OnCorpseOpened - invalid m_curCorpse - Stopping loot"
        Call StopLooting("OnCorpseOpened - Invalid m_curCorpse")
    Else
        locDebug "OnCorpseOpened : " & m_curCorpse.Name & " - Now Searching for Items"
        Call SetSubState(ST_LOOT_SEARCH_ITEM)
        Call m_tmrNextTime.ExpireNow
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.OnCorpseOpened - " & Err.Description
    Resume Fin
End Sub



'Returns true if the item can be assessed
Private Function CanAssess(objItem As acObject)
On Error GoTo ErrorHandler

    Dim bRet As Boolean

    ' If we might have a rare, ID everything!
    If g_bLootRare Then
        bRet = True
        GoTo Fin
    End If
    
    'Only Assess if ID time expired
    'If bRet Then bRet = ((g_ds.Time - objItem.LastIDTime) > ID_EXPIRATION_TIME)
    If IsWorthAssessing(objItem.itemType) Then
        bRet = ((g_ds.Time - objItem.LastIdTime) > ID_EXPIRATION_TIME)
        If Not bRet Then
            locDebug "CanAssess(" & objItem.Name & ") : ID Time Not Expired - Not assessing item"
            locDebug "CanAssess(" & objItem.Name & ") : Time: " & g_ds.Time & " LastIdTime: " & objItem.LastIdTime
        End If
    End If
    
Fin:
    CanAssess = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.CanAssess - " & Err.Description
    Resume Fin
End Function

'Returns true if the item has met one of the filters criteras
Public Function PassedPickupFilters(obj As acObject)
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    
    'Default to no pickup
    bOk = False
    
    If Not Valid(obj) Then
        PrintErrorMessage "clsLoot.PassedPickupFilters - invalid obj - Skipping"
        GoTo Fin
    End If
    
    If g_ui.Loot.ItemIsPickable(obj.Name) Then
        locDebug "PassedPickupFilters(" & obj.Name & ") - OK : Item in Specials List"
        bOk = True
        GoTo Fin
            
    ElseIf MustIdItems And IsWorthAssessing(obj.itemType) Then
    
        'if ID on this item still pending, ignore it until we receive the stats
        If m_ItemsIdQueue.Exists(obj.Guid) Then
            locDebug "PassedPickupFilters: ID still pending on " & obj.Name & " - Returning FALSE"
            bOk = False
            GoTo Fin
        End If
        
        'The filters are evaluated in order of priority
        'As soon as the item passes a filter, the tests stop and the item can be picked up
        'FIXME : may have to still pass through the other tests later on to set special flags
        'such as "is salvage" or to set a "protected" flag on uber armor/weapons pickups
        
        bOk = PhatLoot.PassActiveFilters(obj, _
                                        True, _
                                        True, _
                                        True, _
                                        True, _
                                        g_ui.Loot.chkLootAny.Checked, _
                                        g_ui.Loot.chkLootAny.Checked)
    
    End If

    If Not bOk Then
        locDebug "PassedPickupFilters(" & obj.Name & ") - NO : Item failed to pass any loot filters"
    End If

Fin:
    PassedPickupFilters = bOk
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.PassedPickupFilters - " & Err.Description
    bOk = False
    Resume Fin
End Function


Private Function MustIdItems() As Boolean
    MustIdItems = g_ui.Loot.chkLootAny.Checked _
                    Or g_ui.Loot.chkLootArmors.Checked _
                    Or g_ui.Loot.chkLootWeapons.Checked _
                    Or g_ui.Loot.chkLootWands.Checked _
                    Or g_ui.Loot.chkLootSalvages.Checked
End Function

Public Function ValidPickup(obj As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    bRet = False
    
    locDebug "clsLoot.ValidPickup: " & obj.Name & " type: " & obj.itemType
        
    If Not Valid(obj) Then
        PrintErrorMessage "clsLoot.ValidPickup - invalid obj"
        GoTo Fin
    ElseIf (obj.itemType = ITEM_CORPSE) _
        Or (obj.itemType = ITEM_CONTAINER) _
        Or (obj.itemType = ITEM_ARROW) _
        Or (obj.itemType = ITEM_PORTAL) _
        Or (obj.itemType = ITEM_LIFESTONE) _
        Or (obj.itemType = ITEM_DOOR) Then
        'Or (obj.itemType = ITEM_UNKNOWN) Then
        GoTo Fin
    End If
    
    bRet = PassedPickupFilters(obj)
    
Fin:
    ValidPickup = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.ValidPickup - " & Err.Description
    bRet = False
    Resume Fin
End Function

Private Sub RemoveItemFromCorpse(objItem As acObject)
On Error GoTo ErrorHandler
    
    If Valid(objItem) Then
        Call m_ItemsOnCorpse.Remove(objItem.Guid)  'remove item from list of items on corpse, if present
        Call m_ItemsToLoot.Remove(objItem.Guid)    'remove item from list of items to loot, if present
        locDebug "RemoveItemFromCorpse: removed from corpse: " & objItem.Name
    Else
        PrintErrorMessage "clsLoot.RemoveItemFromCorpse - Invalid objItem"
    End If

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.RemoveItemFromCorpse - " & Err.Description
    Resume Fin
End Sub

Public Sub OnItemPickup(ByVal curItem As acObject)
On Error GoTo Error_Handler
    
    If Not Valid(curItem) Then
        PrintErrorMessage "clsLoot.OnItemPickup - invalid curItem, ignoring"
    Else
        If Valid(m_curItem) Then
            ' Check to see if the Item we just looted was the item we ment to pick up
            If m_curItem.Guid = curItem.Guid Then
                MyDebug "OnItemPickup - " & curItem.Name & " picked up."
                If curItem.LastIdTime < g_ds.Time Then
                    Call g_Hooks.IDQueueAdd(curItem.Guid) 'Make sure we ID it
                End If

                RaiseEvent OnItemLooted(curItem)
                Call RemoveItemFromCorpse(curItem)
                Call m_tmrNextTime.ExpireNow
                Call SetSubState(ST_LOOT_ITEM_PICKED_UP)
                
                If PhatLoot.PassActiveFilters(curItem, True, False, False, False, False, False) And g_Macro.Salvager.CheckValidSalvageItem(curItem) Then
                    MyDebug "clsLoot: OnItemPickup: Adding to SalvageList: " & curItem.Name
                    Call g_Macro.Salvager.AddLootToSalvageList(curItem)
                'See if we should add it to the Mana Stone charger list
                ElseIf PhatLoot.CheckHighManaItem(curItem) Then
                    'Add to High Mana item collection
                    If m_colHighManaItems.addObject(curItem) Then
                        MyDebug "clsLoot.OnPickUp: added to HighManaItems collection: " & curItem.Name
                    Else
                        MyDebug "clsLoot.OnPickUp: colHighManaItems.addObject FAILED: " & curItem.Name
                    End If
                Else
                    locDebug "clsLoot: OnItemPickup: did NOT pass salvage Filters: " & curItem.Name
                    locDebug "clsLoot: PhatLoot: " & PhatLoot.PassActiveFilters(curItem, True, False, False, False, False, False)
                    locDebug "clsLoot: ValidSalvage: " & g_Macro.Salvager.CheckValidSalvageItem(curItem)
                End If
                
                If (curItem.itemType = ITEM_MANA_STONES) And (curItem.Name Like "*Stone*") And (curItem.Mana = 0) Then
                    Call g_Macro.Loot.countEmptyManaStones
                End If

            End If
        End If
    End If
    
Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsLoot.OnItemPickup - " & Err.Description
    Resume Fin
End Sub

'Looks up for items we can pickup on the current body
'PreConditions : m_curCorpse not null, m_curCorpse within range
Private Function FoundItemToPickup(ByRef objItemOut As acObject) As Boolean
On Error GoTo Error_Handler

    Dim bRet As Boolean
    Dim tmpObj As acObject

    'default ret val
    bRet = False
    Set objItemOut = Nothing
    
    If m_ItemsToLoot.Count > 0 Then
        locDebug "FoundItemToPickup - Current Loot List : " & m_ItemsToLoot.GetObjectsList
        
        'Get the first item we find in the collection
        For Each tmpObj In m_ItemsToLoot
            'update current item
            Set objItemOut = tmpObj
            
            'return true
            bRet = True
            GoTo Fin
        Next tmpObj
    End If
    
Fin:
    FoundItemToPickup = bRet
    Set tmpObj = Nothing
    Exit Function
Error_Handler:
    PrintErrorMessage "clsLoot.FoundItemToPickup - " & Err.Description
    Resume Fin
End Function

Private Function DoOpenCorpse() As Boolean
On Error GoTo Error_Handler
    locDebug "clsLoot.DoOpenCorpse: Opening corpse..."
    
    Dim iCount As Integer
    Dim bRet As Boolean
    
    If (m_curCorpse.GetSquareRange = 0) Then
        'Range of 0 is an error
        Call m_curCorpse.SetUserData(B_LOOTED, True)
        Call StopLooting("DoOpen Corpse - GetSquareRange is ZERO")
    ElseIf (m_curCorpse Is Nothing) _
     Or (m_curCorpse.GetSquareRange > (g_Data.LootSearchRadius * g_Data.LootSearchRadius)) _
     Or (m_curCorpse.UserData(B_LOOTED)) _
     Or (Not CanLootCorpse(m_curCorpse)) Then
        bRet = False
        locDebug "DoOpenCorpse: invalid corpse."
        
        If (m_curCorpse Is Nothing) Then MyDebug "DoOpenCorpse: m_curCorpse Is Nothing"
        If (m_curCorpse.GetSquareRange > (g_Data.LootSearchRadius * g_Data.LootSearchRadius)) Then MyDebug "DoOpenCorpse: out of range"
        If (m_curCorpse.UserData(B_LOOTED)) Then MyDebug "DoOpenCorpse: UserData(B_LOOTED) is true"
        
        Set m_curCorpse = Nothing
    Else
        bRet = True
        Call g_Service.UseItem(m_curCorpse)    'UseItemOnSelf
        Call m_tmrNextTime.SetNextTime(3)      'Next corpse open try
        iCount = m_curCorpse.UserData(INT_BLISTCOUNT) + 1
        Call m_curCorpse.SetUserData(INT_BLISTCOUNT, iCount)
    End If
    
Fin:
    DoOpenCorpse = bRet
    Exit Function
Error_Handler:
    PrintErrorMessage "clsLoot.DoOpenCorpse : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    bRet = False
    Resume Fin
End Function

Private Function DoPickupItem() As Boolean
On Error GoTo Error_Handler
    
    If Not Valid(m_curItem) Then
        PrintErrorMessage "clsLoot.DoPickupItem: invalid m_curItem, skipping"
        DoPickupItem = False
        Set m_curItem = Nothing
    ElseIf g_Objects.Items.InInventory(m_curItem.Guid) Then
        locDebug "DoPickupItem - " & m_curItem.Name & " already in inventory, ignoring pickup request."
        'Call OnItemPickup(m_curItem)
        DoPickupItem = True
    Else
        If m_bGroundPickup Then
            'PrintMessage "Looting item on ground : " & m_curItem.Name
            locDebug "DoPickupItem - picking up ground item : " & m_curItem.Name
        Else
            locDebug "DoPickupItem - picking up " & m_curItem.Name & " on " & m_curCorpse.Name
        End If
        
        Call g_Service.UseItem(m_curItem) 'UseItemOnSelf

        Call m_tmrNextTime.SetNextTime(1)
        DoPickupItem = True
    End If
    
Fin:
    Exit Function
Error_Handler:
    PrintErrorMessage "clsLoot.DoPickupItem : Error #" & Err.Number & " (line: " & Erl & ") has been generated by " & Err.Source & " : " & Err.Description
    Resume Fin
    
End Function

Private Sub SetBodyLooted(Optional ByVal bLooted As Boolean = True)
On Error GoTo ErrorHandler

    If Valid(m_curCorpse) Then
        Call m_curCorpse.SetUserData(B_LOOTED, bLooted)
        g_TotalLooted = g_TotalLooted + 1
        MyDebug "clsLoot.SetBodyLooted:  TotalKilled: " & g_TotalKilled & "  TotalLooted: " & g_TotalLooted
        
        'Decrement the waypoint hold.
        Call decNavHold
        
        If g_bLootRare = True Then
            MyDebug "A rare still needs to be looted!"
        End If
        
        If m_curCorpse.UserData(B_HASRARE) Then
            'Make sure to unset rare stuff
            Set m_rareCorpse = Nothing
            g_bLootRare = False
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.SetBodyLooted - " & Err.Description
    Resume Fin
End Sub

Public Sub IgnoreCurrentCorpse()
    If Not Valid(m_curCorpse) Then Exit Sub
    Call AddCorpseToIgnore(m_curCorpse)
End Sub

'Bodies to Ignore
'GUID -> Time when added to the list
Private Function AddCorpseToIgnore(ByVal objCorpse As acObject) As Boolean
On Error GoTo ErrorHandler
    
    If Not Valid(objCorpse) Then GoTo Fin
    
    If m_dicCorpseToIgnore.Exists(objCorpse.Guid) Then
        'Update time
        m_dicCorpseToIgnore(objCorpse.Guid) = g_Core.Time
        locDebug "AddCorpseToIgnore - Updated timer on " & objCorpse.Name
    Else
        Call m_dicCorpseToIgnore.Add(objCorpse.Guid, g_Core.Time)
        locDebug "AddCorpseToIgnore - Added " & objCorpse.Name
        Call DebugIgnoreCorpseList
    End If
    
    AddCorpseToIgnore = True
    
Fin:
    Exit Function
ErrorHandler:
    AddCorpseToIgnore = False
    PrintErrorMessage "clsLoot.AddCorpseToIgnore - " & Err.Description
    Resume Fin
End Function

Private Function RemoveCorpseToIgnore(ByVal objCorpse As acObject) As Boolean
On Error GoTo ErrorHandler
    
    If m_dicCorpseToIgnore.Exists(objCorpse.Guid) Then
        Call m_dicCorpseToIgnore.Remove(objCorpse.Guid)
        locDebug "RemoveCorpseToIgnore - Removed " & objCorpse.Name
        RemoveCorpseToIgnore = True
    End If
    
    RemoveCorpseToIgnore = False
    
Fin:
    Exit Function
ErrorHandler:
    RemoveCorpseToIgnore = False
    PrintErrorMessage "clsLoot.RemoveCorpseToIgnore - " & Err.Description
    Resume Fin
End Function

Private Function MustIgnoreCorpse(ByVal objCorpse As acObject, Optional fExpirationTime As Double = 15) As Boolean
On Error GoTo ErrorHandler
    
    If Not Valid(objCorpse) Then
        ' clean out the entire list
        ' FIXME -- this should probably be on a 5 min timer or something
         m_dicCorpseToIgnore.RemoveAll
        locDebug "clsLoot.MustIgnoreCorpse - removed ALL objects "
        GoTo Fin
    End If
    
    If m_dicCorpseToIgnore.Exists(objCorpse.Guid) Then
        Dim fAddTime As Double, fDelta As Double
        fAddTime = m_dicCorpseToIgnore(objCorpse.Guid)
        fDelta = g_Core.Time - fAddTime
        'locDebug "MustIgnoreCorpse - " & objCorpse.Name & " - fDelta: " & fDelta & " sec left"
        If fDelta < fExpirationTime Then
            'locDebug "MustIgnoreCorpse - " & objCorpse.Name & " time isnt up yet - Returning true"
            MustIgnoreCorpse = True
            GoTo Fin
        Else
            'Delay time is up, remove it from list
            'locDebug "MustIgnoreCorpse - Blacklist time up, removing corpse from ignore list"
            Call RemoveCorpseToIgnore(objCorpse)
        End If
    End If
    
    MustIgnoreCorpse = False
    
Fin:
    Exit Function
ErrorHandler:
    MustIgnoreCorpse = False
    PrintErrorMessage "clsLoot.MustIgnoreCorpse - " & Err.Description
    Resume Fin
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Local Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function GetStateString() As String
Dim sRes As String

    Select Case m_State
        Case ST_LOOT_SEARCH
            sRes = "Searching Corpse"
            
        Case ST_LOOT_OPEN_CORPSE
            sRes = "Opening Corpse"
            
        Case ST_LOOT_SEARCH_ITEM
            sRes = "Searching Items on Corpse"
            
        Case ST_LOOT_PICKUP_ITEM
            sRes = "Picking up Item"
            
        Case ST_LOOT_ITEM_PICKED_UP
            sRes = "Item Picked Up"
            
        Case ST_LOOT_STOP
            sRes = "Stopping Looting"
            
        Case Else
            sRes = "Unknown State " & m_State
    End Select
    
    GetStateString = sRes
End Function

' Print out an ID string
Private Sub printIdString(ByVal aObj As acObject)
    Dim anID As String
    anID = getIdString(aObj)
    If Len(anID) > 1 Then PrintMessage anID
End Sub

Private Function getIdString(ByVal aObj As acObject) As String
On Error GoTo ErrorHandler

    Dim id As String
    Dim bPrint As Boolean
    Dim i As Integer
    Dim aVar As Variant
    
    bPrint = False

    If Not Valid(aObj) Then GoTo Fin

    id = aObj.Name
    
    Select Case aObj.itemType
        Case ITEM_ARMOR
            'Brass Platemail Greaves, 6 tinks, (7) craft, AL 350(370),
            ' Major Coordination, Impen VI, Acid Bane VI, Minor Acid Bane,
            ' Diff. 135, Melee 265, ThananimBlack, Value 8,336, 642BU
            ' -->
            ' Greaves, 6 tinks, work 7, AL 350, <Spells>, Lore 135, Melee 265, Value 8336, 642 BU
            
            If (aObj.TinkCount > 0) Then id = id & ", " & aObj.TinkCount & " tinks"
            id = id & ", work " & aObj.Workmanship
            id = id & ", AL " & getRealArmorLevel(aObj)
            id = id & getUberArmorRating(aObj)
            
            bPrint = True
            
        Case ITEM_MELEE_WEAPON
            MyDebug "Found a melee weapon."
            'aName, 24.6-41 +11m/+13a, <spells>, Diff. 200, UA 370+ to wield, Value 6,000, 110BU
            'If (aObj.TinkCount > 0) Then id = id & ", " & aObj.TinkCount & " tinks"
            'If (aObj.TinkCount > 0) Then
                If (Len(aObj.Imbue) > 1) Then id = id & ", (" & aObj.Imbue
                
                'If aObj.CrushingBlow = True And (Len(aObj.Imbue) > 1) Then
                '    id = id & ", Crushing Blow"
                'ElseIf aObj.CrushingBlow = True And (Len(aObj.Imbue) < 1) Then
                '    'MyDebug "Crushing Blow found and no imbue string."
                '    id = id & ", (Crushing Blow"
                'End If
                
                'If aObj.BitingStrike = True And (Len(aObj.Imbue) > 1 Or aObj.CrushingBlow = True) Then
                '    id = id & ", Biting Strike"
                'ElseIf aObj.BitingStrike = True And (Len(aObj.Imbue) < 1 Or aObj.CrushingBlow = False) Then
                '    'MyDebug "Biting Strike found and no imbue string."
                '    If aObj.CrushingBlow = False Then
                '        id = id & ", (Biting Strike"
                '    End If
                'End If

                If (Len(aObj.ResistanceCleaving) > 1) Then
                    If ((Len(aObj.ResistanceCleaving) > 1) Or aObj.CrushingBlow = True Or aObj.BitingStrike = True) Then
                        id = id & ", " & aObj.ResistanceCleaving & aObj.ResistanceCleavingType & ")"
                    Else
                        id = id & "(" & aObj.ResistanceCleaving & aObj.ResistanceCleavingType & ")"
                    End If
                ElseIf (Len(aObj.Imbue) > 1) Then
                    id = id & ")"
                ElseIf aObj.BitingStrike = True Then
                    id = id & ")"
                ElseIf aObj.CrushingBlow = True Then
                    id = id & ")"
                End If
                
                id = id & ", " & aObj.TinkCount & " tinks"
            'End If
            id = id & ", work " & aObj.Workmanship
            id = id & ", dmg " & getRealMeleeDamage(aObj, False) & "-" & getRealMeleeDamage(aObj, True)
            id = id & getUberMeleeRating(aObj)
            id = id & ", +" & getRealAttackBonus(aObj) & " atk"
            If (aObj.DefenseBonus > 0) Then id = id & "/+" & getRealMeleeD(aObj) & " MeleeD"
            If (aObj.MagicDefense > 0) Then id = id & "/+" & aObj.MagicDefense & " MagicD"
            If (aObj.MissileDefense > 0) Then id = id & "/+" & aObj.MissileDefense & " MissileD"
            If (aObj.ElementBonusDamage > 0) Then id = id & "/+" & aObj.ElementBonusDamage & " DMG"
            If (aObj.WieldReqVal > 0) And (aObj.WieldReqType = 2) Then
                id = id & ", " & g_ACConst.GetWeaponSkillName(aObj.WieldReqID)
                id = id & " " & aObj.WieldReqVal & "+ to wield"
            ElseIf (aObj.WieldReqVal > 0) And (aObj.WieldReqType = 7) Then
                id = id & ", " & g_ACConst.GetLevelName(aObj.WieldReqID)
                id = id & " " & aObj.WieldReqVal & "+ to wield"
            End If
            
            bPrint = True
            
        Case ITEM_MISSILE_WEAPON
            MyDebug "Found a missile weapon."
            'FIXME
            'Ruby Chalice, (8) craft, +0% +0m, Value 24,365, 50BU
            'If (aObj.TinkCount > 0) Then id = id & ", " & aObj.TinkCount & " tinks"
            'If (aObj.TinkCount > 0) Then
                If (Len(aObj.Imbue) > 1) Then id = id & ", (" & aObj.Imbue
                
                'If aObj.CrushingBlow = True And Len(aObj.Imbue) > 1 Then
                '    id = id & ", Crushing Blow"
                'ElseIf aObj.CrushingBlow = True And (Len(aObj.Imbue) < 1) Then
                '    'MyDebug "Crushing Blow found and no imbue string."
                '    id = id & ", (Crushing Blow"
                'End If
                
                'If aObj.BitingStrike = True And (Len(aObj.Imbue) > 1 Or aObj.CrushingBlow = True) Then
                '    id = id & ", Biting Strike"
                'ElseIf aObj.BitingStrike = True And (Len(aObj.Imbue) < 1 Or aObj.CrushingBlow = False) Then
                '    'MyDebug "Biting Strike found and no imbue string."
                '    If aObj.CrushingBlow = False Then
                '        id = id & ", (Biting Strike"
                '    End If
                'End If

                If (Len(aObj.ResistanceCleaving) > 1) Then
                    id = id & aObj.ResistanceCleaving & aObj.ResistanceCleavingType & ")"
                ElseIf (Len(aObj.Imbue) > 1) Then
                    id = id & ")"
                ElseIf aObj.BitingStrike = True Then
                    id = id & ")"
                ElseIf aObj.CrushingBlow = True Then
                    id = id & ")"
                End If
                
                id = id & ", " & aObj.TinkCount & " tinks"
            'End If
            id = id & ", work " & aObj.Workmanship
            id = id & ", +" & aObj.DamageModifier & "% "
            id = id & getUberRangedRating(aObj)
            id = id & " +" & getRealMeleeD(aObj) & " MeleeD"
            If (aObj.MagicDefense > 0) Then id = id & "/+" & aObj.MagicDefense & " MagicD"
            If (aObj.MissileDefense > 0) Then id = id & "/+" & aObj.MissileDefense & " MissileD"
            If (aObj.ElementBonusDamage > 0) Then id = id & "/+" & aObj.ElementBonusDamage & " DMG"
            If (aObj.WieldReqVal > 0) And (aObj.WieldReqType = 2) Then
                id = id & ", " & g_ACConst.GetWeaponSkillName(aObj.WieldReqID)
                id = id & " " & aObj.WieldReqVal & "+ to wield"
            ElseIf (aObj.WieldReqVal > 0) And (aObj.WieldReqType = 7) Then
                id = id & ", " & g_ACConst.GetLevelName(aObj.WieldReqID)
                id = id & " " & aObj.WieldReqVal & "+ to wield"
            End If
            
            bPrint = True
            
        Case ITEM_WAND
            MyDebug "Found a magic weapon."
            'Ruby Staff, 6 tinks, (7) craft, +9mc/+15m,
            'Major Item Enchantment Aptitude, War Self VI, Diff. 248, Value 22,593, 50BU
            
            'If (aObj.TinkCount > 0) Then
                If (Len(aObj.Imbue) > 1) Then id = id & ", (" & aObj.Imbue
                
                'If aObj.CrushingBlow = True And Len(aObj.Imbue) > 1 Then
                '    id = id & ", Crushing Blow"
                'ElseIf aObj.CrushingBlow = True And (Len(aObj.Imbue) < 1) Then
                '    'MyDebug "Crushing Blow found and no imbue string."
                '    id = id & ", (Crushing Blow"
                'End If
                
                'If aObj.BitingStrike = True And (Len(aObj.Imbue) > 1 Or aObj.CrushingBlow = True) Then
                '    id = id & ", Biting Strike"
                'ElseIf aObj.BitingStrike = True And (Len(aObj.Imbue) < 1 Or aObj.CrushingBlow = False) Then
                '    'MyDebug "Biting Strike found and no imbue string."
                '    If aObj.CrushingBlow = False Then
                '        id = id & ", (Biting Strike"
                '    End If
                'End If

                If (Len(aObj.ResistanceCleaving) > 1) Then
                    id = id & ", " & aObj.ResistanceCleaving & aObj.ResistanceCleavingType & ")"
                ElseIf (Len(aObj.Imbue) > 1) Then
                    id = id & ")"
                ElseIf aObj.BitingStrike = True Then
                    id = id & ")"
                ElseIf aObj.CrushingBlow = True Then
                    id = id & ")"
                End If
                
                id = id & ", " & aObj.TinkCount & " tinks"
            'End If
            id = id & ", work " & aObj.Workmanship
            If (aObj.ManaConvMod > 0) Then id = id & ", +" & getRealManaC(aObj) & " ManaC"
            If (aObj.DefenseBonus > 0) Then id = id & "/+" & getRealMeleeD(aObj) & " MeleeD"
            If (aObj.MagicDefense > 0) Then id = id & "/+" & aObj.MagicDefense & " MagicD"
            If (aObj.MissileDefense > 0) Then id = id & "/+" & aObj.MissileDefense & " MissileD"
            If (aObj.PvMBonus > 0) Then id = id & "/+" & getRealPvM(aObj) & " PvM"
            
            bPrint = True
        
        Case Else
            'Only output if it's something interesting, like a major
            If aObj.HasMajors Or aObj.HasMinors Or aObj.IsRare Or (aObj.ArmorLevel > 0) Then
                
                If (Len(aObj.Imbue) > 1) Then id = id & ", (" & aObj.Imbue
                
                'If aObj.CrushingBlow = True And Len(aObj.Imbue) > 1 Then
                '    id = id & ", Crushing Blow"
                'ElseIf aObj.CrushingBlow = True And (Len(aObj.Imbue) < 1) Then
                '    MyDebug "Crushing Blow found and no imbue string."
                '    id = id & ", (Crushing Blow"
                'End If
                
                'If aObj.BitingStrike = True And (Len(aObj.Imbue) > 1 Or aObj.CrushingBlow = True) Then
                '    id = id & ", Biting Strike"
                'ElseIf aObj.BitingStrike = True And (Len(aObj.Imbue) < 1 Or aObj.CrushingBlow = False) Then
                '    MyDebug "Biting Strike found and no imbue string."
                '    If aObj.CrushingBlow = False Then
                '        id = id & ", (Biting Strike"
                '    End If
                'End If

                If (Len(aObj.ResistanceCleaving) > 1) Then
                    id = id & aObj.ResistanceCleaving & aObj.ResistanceCleavingType & ")"
                ElseIf (Len(aObj.Imbue) > 1) Then
                    id = id & ")"
                ElseIf aObj.BitingStrike = True Then
                    id = id & ")"
                ElseIf aObj.CrushingBlow = True Then
                    id = id & ")"
                End If
                
                If (aObj.TinkCount > 0) Then id = id & ", " & aObj.TinkCount & " tinks"
                If (aObj.Workmanship > 0) Then id = id & ", work " & aObj.Workmanship
                If (aObj.ArmorLevel > 0) Then
                    id = id & ", AL " & getRealArmorLevel(aObj)
                    id = id & getUberArmorRating(aObj)
                End If
                
                bPrint = True
            End If
    
    End Select

    If bPrint Then
    
        'Check for innate spells on this item
        If Valid(aObj.Spells) Then
            If (aObj.Spells.Count > 0) Then
                For Each aVar In aObj.Spells
                    If (aObj.Spells.Item(aVar) = 0) Then
                        id = id & ", " & CStr(aVar)
                    End If
                Next
            End If
        End If
        
        If (aObj.LoreReq > 0) Then id = id & ", Lore " & aObj.LoreReq
        If (aObj.RankReq > 0) Then id = id & ", Rank " & aObj.RankReq
        If (aObj.ActivateSkillVal > 0) Then id = id & ", " & aObj.ActivateSkill & " " & aObj.ActivateSkillVal
        
        'Wield Req ID #2 is ATTRIBUTE required to wield (Sword, Axe, Bow, etc)
        If (((aObj.WieldReqVal > 0) And (aObj.WieldReqType = 2)) And ((aObj.itemType = ITEM_ARMOR) Or (aObj.itemType = ITEM_WAND))) Then
            id = id & ", " & g_ACConst.GetSkillName(aObj.WieldReqID)
            id = id & " " & aObj.WieldReqVal & "+ to wield"
        End If
        
        'Wield Req ID #7 is the TYPE required to wield (Level, so far only one found) Added for version 11.7.0
        If (((aObj.WieldReqVal > 0) And (aObj.WieldReqType = 7)) And ((aObj.itemType = ITEM_ARMOR) Or (aObj.itemType = ITEM_WAND))) Then
            id = id & ", " & g_ACConst.GetLevelName(aObj.WieldReqID)
            id = id & " " & aObj.WieldReqVal & "+ to wield"
        End If
        
        'Wield Req ID #9 is for new Societies
        If (((aObj.WieldReqVal > 0) And (aObj.WieldReqType = 9)) And ((aObj.itemType = ITEM_ARMOR) Or (aObj.itemType = ITEM_WAND))) Then
            id = id & ", Society Rank"
            id = id & " " & aObj.WieldReqVal & "+ to wield"
        End If
        
        'If ((aObj.itemType = ITEM_ARMOR) Or (aObj.itemType = ITEM_WAND) Or (aObj.itemType = ITEM_MELEE_WEAPON) Or (aObj.itemType = ITEM_MISSILE_WEAPON)) And (aObj.WieldReqVal > 0) Then
        '    id = id & ", " & g_ACConst.GetSkillName(aObj.WieldReqID)
        '    id = id & " " & aObj.WieldReqVal & "+ to wield"
        'End If
        
        If (Len(aObj.RaceReq) > 1) Then id = id & ", " & aObj.RaceReq
        id = id & ", Value " & aObj.Value
        id = id & ", " & aObj.Burden & " BU"
                
        getIdString = id
        MyDebug getIdString
        
    End If


Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getIdString - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

'            If (aObj.ManaConvMod > 0) Then id = id & ", +" & aObj.ManaConvMod & " ManaC"
'            If (aObj.DefenseBonus > 0) Then id = id & "/+" & aObj.DefenseBonus & " MeleeD"
'            If (aObj.PvMBonus > 0) Then id = id & "/+" & aObj.PvMBonus & " PvM" ' Spirit Drinker

Public Function getRealPvM(ByVal aObj As acObject) As Integer
On Error GoTo ErrorHandler

    Dim realPvM As Integer
    Dim buffLevel As Integer
    Dim aVar As Variant

    realPvM = aObj.PvMBonus
    buffLevel = 0
    
    'Check for active cast spells on this item
    If Valid(aObj.SpellsActive) Then
        If (aObj.SpellsActive.Count > 0) Then
        
            For Each aVar In aObj.SpellsActive
                    Dim buffSpell As clsSpell
                        
                    Set buffSpell = g_Spells.Items.FindSpellByName(CStr(aVar))
                    If Valid(buffSpell) Then
                        'Need to recalc Armor value based on active spells
                        If buffSpell.SpellFamily = "Spirit Drinker" Then
                            MyDebug "clsLoot.getRealManaC: " & aObj.Name & " has Spirit Drinker active: " & buffSpell.SpellName
                            If buffLevel < buffSpell.SpellLevel Then buffLevel = buffSpell.SpellLevel
                        End If
                    End If

            Next
        End If
    End If

    If buffLevel = 1 Then
        realPvM = realPvM - 1
    ElseIf buffLevel = 2 Then
        realPvM = realPvM - 2
    ElseIf buffLevel = 3 Then
        realPvM = realPvM - 3
    ElseIf buffLevel = 4 Then
        realPvM = realPvM - 4
    ElseIf buffLevel = 5 Then
        realPvM = realPvM - 5
    ElseIf buffLevel = 6 Then
        realPvM = realPvM - 6
    ElseIf buffLevel = 7 Then
        realPvM = realPvM - 7
    End If

Fin:
    getRealPvM = realPvM
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getRealPvM - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function getRealMeleeD(ByVal aObj As acObject) As Integer
On Error GoTo ErrorHandler

    Dim realMeleeD As Integer
    Dim buffLevel As Integer
    Dim aVar As Variant

    realMeleeD = aObj.DefenseBonus
    buffLevel = 0
    
    'Check for active cast spells on this item
    If Valid(aObj.SpellsActive) Then
        If (aObj.SpellsActive.Count > 0) Then
        
            For Each aVar In aObj.SpellsActive
                    Dim buffSpell As clsSpell
                        
                    Set buffSpell = g_Spells.Items.FindSpellByName(CStr(aVar))
                    If Valid(buffSpell) Then
                        'Need to recalc Armor value based on active spells
                        If buffSpell.SpellFamily = "Defender" Then
                            MyDebug "clsLoot.getRealMeleeD: " & aObj.Name & " has Defender active: " & buffSpell.SpellName
                            If buffLevel < buffSpell.SpellLevel Then buffLevel = buffSpell.SpellLevel
                        End If
                    End If

            Next
        End If
    End If

    If buffLevel = 1 Then
        '2.5%
        realMeleeD = realMeleeD - 2.5
    ElseIf buffLevel = 2 Then
        '5%
        realMeleeD = realMeleeD - 5
    ElseIf buffLevel = 3 Then
        '7.5%
        realMeleeD = realMeleeD - 7.5
    ElseIf buffLevel = 4 Then
        '10%
        realMeleeD = realMeleeD - 10
    ElseIf buffLevel = 5 Then
        '12.5%
        realMeleeD = realMeleeD - 12.5
    ElseIf buffLevel = 6 Then
        '15%
        realMeleeD = realMeleeD - 15
    ElseIf buffLevel = 7 Then
        '17%
        realMeleeD = realMeleeD - 17
    End If

Fin:
    getRealMeleeD = realMeleeD
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getRealMeleeD - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function


Public Function getRealAttackBonus(ByVal aObj As acObject) As Integer
On Error GoTo ErrorHandler

    Dim realAttack As Integer
    Dim buffLevel As Integer
    Dim aVar As Variant

    realAttack = aObj.AttackBonus
    buffLevel = 0
    
    'Check for active cast spells on this item
    If Valid(aObj.SpellsActive) Then
        If (aObj.SpellsActive.Count > 0) Then
        
            For Each aVar In aObj.SpellsActive
                    Dim buffSpell As clsSpell
                        
                    Set buffSpell = g_Spells.Items.FindSpellByName(CStr(aVar))
                    If Valid(buffSpell) Then
                        'Need to recalc Armor value based on active spells
                        If buffSpell.SpellFamily = "Heart Seeker" Then
                            MyDebug "clsLoot.getRealAttackBonus: " & aObj.Name & " has Heart Seeker active: " & buffSpell.SpellName
                            If buffLevel < buffSpell.SpellLevel Then buffLevel = buffSpell.SpellLevel
                        End If
                    End If

            Next
        End If
    End If

    If buffLevel = 1 Then
        '2.5%
        realAttack = realAttack - 2.5
    ElseIf buffLevel = 2 Then
        '5%
        realAttack = realAttack - 5
    ElseIf buffLevel = 3 Then
        '7.5%
        realAttack = realAttack - 7.5
    ElseIf buffLevel = 4 Then
        '10%
        realAttack = realAttack - 10
    ElseIf buffLevel = 5 Then
        '12.5%
        realAttack = realAttack - 12.5
    ElseIf buffLevel = 6 Then
        '15%
        realAttack = realAttack - 15
    ElseIf buffLevel = 7 Then
        '17%
        realAttack = realAttack - 17
    End If

Fin:
    getRealAttackBonus = realAttack
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getRealAttackBonus - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function getRealManaC(ByVal aObj As acObject) As Integer
On Error GoTo ErrorHandler

    Dim realManaC As Integer
    Dim buffLevel As Integer
    Dim aVar As Variant

    realManaC = aObj.ManaConvMod
    buffLevel = 0
    
    'Check for active cast spells on this item
    If Valid(aObj.SpellsActive) Then
        If (aObj.SpellsActive.Count > 0) Then
        
            For Each aVar In aObj.SpellsActive
                    Dim buffSpell As clsSpell
                        
                    Set buffSpell = g_Spells.Items.FindSpellByName(CStr(aVar))
                    If Valid(buffSpell) Then
                        'Need to recalc Armor value based on active spells
                        If buffSpell.SpellFamily = "Hermetic Link" Then
                            MyDebug "clsLoot.getRealManaC: " & aObj.Name & " has Hermetic Link active: " & buffSpell.SpellName
                            If buffLevel < buffSpell.SpellLevel Then buffLevel = buffSpell.SpellLevel
                        End If
                    End If

            Next
        End If
    End If

    If buffLevel = 1 Then
        realManaC = realManaC / 1.1
    ElseIf buffLevel = 2 Then
        realManaC = realManaC / 1.2
    ElseIf buffLevel = 3 Then
        realManaC = realManaC / 1.3
    ElseIf buffLevel = 4 Then
        realManaC = realManaC / 1.4
    ElseIf buffLevel = 5 Then
        realManaC = realManaC / 1.5
    ElseIf buffLevel = 6 Then
        realManaC = realManaC / 1.6
    ElseIf buffLevel = 7 Then
        realManaC = realManaC / 1.7
    End If

Fin:
    getRealManaC = realManaC
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getRealManaC - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function getRealArmorLevel(ByVal aObj As acObject) As Integer
On Error GoTo ErrorHandler

    Dim realAL As Integer
    Dim buffLevel As Integer
    Dim aVar As Variant

    realAL = aObj.ArmorLevel
    buffLevel = 0
    
    'Check for active cast spells on this item
    If Valid(aObj.SpellsActive) Then
        If (aObj.SpellsActive.Count > 0) Then
        
            For Each aVar In aObj.SpellsActive
                    Dim buffSpell As clsSpell
                        
                    Set buffSpell = g_Spells.Items.FindSpellByName(CStr(aVar))
                    If Valid(buffSpell) Then
                        'Need to recalc Armor value based on active spells
                        If buffSpell.SpellFamily = "Impenetrability" Then
                            MyDebug "clsLoot.getRealArmorLevel: " & aObj.Name & " has Impenetrability active: " & buffSpell.SpellName
                            If buffLevel < buffSpell.SpellLevel Then buffLevel = buffSpell.SpellLevel
                        End If
                    End If

            Next
        End If
    End If

    If buffLevel = 1 Then
        '<lvl=1> by 20 points.>
        realAL = aObj.ArmorLevel - 20
    ElseIf buffLevel = 2 Then
        '<lvl=2> by 50 points.>
        realAL = aObj.ArmorLevel - 50
    ElseIf buffLevel = 3 Then
        '<lvl=3> by 75 points.>
        realAL = aObj.ArmorLevel - 75
    ElseIf buffLevel = 4 Then
        '<lvl=4> by 100 points.>
        realAL = aObj.ArmorLevel - 100
    ElseIf buffLevel = 5 Then
        '<lvl=5> by 150 points.>
        realAL = aObj.ArmorLevel - 150
    ElseIf buffLevel = 6 Then
        '<lvl=6> by 200 points.>
        realAL = aObj.ArmorLevel - 200
    ElseIf buffLevel = 7 Then
        '<lvl=7> by 220 points.>
        realAL = aObj.ArmorLevel - 220
    End If

Fin:
    getRealArmorLevel = realAL
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getRealArmorLevel - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function getRealMeleeDamage(ByVal aObj As acObject, ByVal getHigh As Boolean) As Integer
On Error GoTo ErrorHandler

    Dim realDamage As Integer
    Dim buffLevel As Integer
    Dim aVar As Variant

    If getHigh Then
        realDamage = aObj.HighDamage
        MyDebug "clsLoot.getRealMeleeDamage: realDamage: " & realDamage & " :High: " & aObj.HighDamage
    Else
        realDamage = Round(aObj.HighDamage - (aObj.Variance * aObj.HighDamage), 2)
        'realDamage = aObj.GetLowDamage
        MyDebug "clsLoot.getRealMeleeDamage: Variance: " & aObj.Variance & " :High: " & aObj.HighDamage
        MyDebug "clsLoot.getRealMeleeDamage: realDamage: " & realDamage & " :Low: " & aObj.GetLowDamage
    End If
    
    realDamage = aObj.HighDamage
    buffLevel = 0
    
    'Check for active cast spells on this item
    If Valid(aObj.SpellsActive) Then
        If (aObj.SpellsActive.Count > 0) Then
        
            For Each aVar In aObj.SpellsActive
                    Dim buffSpell As clsSpell
                        
                    Set buffSpell = g_Spells.Items.FindSpellByName(CStr(aVar))
                    If Valid(buffSpell) Then
                        'Need to recalc Damage value based on active spells
                        If buffSpell.SpellFamily = "Blood Drinker" Then
                            MyDebug "clsLoot.getRealMeleeDamage: " & aObj.Name & " has Blood Drinker active: " & buffSpell.SpellName
                            If buffLevel < buffSpell.SpellLevel Then buffLevel = buffSpell.SpellLevel
                        End If
                    End If

            Next
        End If
    End If

    If buffLevel = 1 Then
        realDamage = realDamage - 2
    ElseIf buffLevel = 2 Then
        realDamage = realDamage - 4
    ElseIf buffLevel = 3 Then
        realDamage = realDamage - 8
    ElseIf buffLevel = 4 Then
        realDamage = realDamage - 12
    ElseIf buffLevel = 5 Then
        realDamage = realDamage - 16
    ElseIf buffLevel = 6 Then
        realDamage = realDamage - 20
    ElseIf buffLevel = 7 Then
        realDamage = realDamage - 22
    End If

    If Not getHigh Then
        realDamage = Round(realDamage - (aObj.Variance * realDamage), 2)
        MyDebug "clsLoot.getRealmeleeDamage: Variance: " & aObj.Variance & " :High: " & aObj.HighDamage
        MyDebug "clsLoot.getRealmeleeDamage: realDamage: " & realDamage & " :Low: " & aObj.GetLowDamage
    End If

Fin:
    getRealMeleeDamage = realDamage
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getRealMeleeDamage - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function getUberRangedRating(ByVal aObj As acObject) As String
On Error GoTo ErrorHandler

    Dim id As String
    Dim wieldReq, eDamage, damageMod As Integer

    wieldReq = 0
    eDamage = 0
    damageMod = 0
    
    MyDebug "getUberRangedRating: " & aObj.Name & " WieldType: " & aObj.WieldReqType & " (" & aObj.WieldReqVal & ") :: " & g_ACConst.GetWeaponSkillName(aObj.WieldReqID)
    MyDebug "getUberRangedRating: Dmg: " & aObj.DamageModifier & "%  -- eDamage: " & aObj.ElementBonusDamage


    If (aObj.WieldReqVal > 0) And (aObj.WieldReqType = 2) Then
        Select Case g_ACConst.GetWeaponSkillName(aObj.WieldReqID)
            Case "Bow"
                ' No      250     270     290     315         335     Bonus   360     Bonus
                ' 100%    110%    120%    130%    130% (+4)   130%    +5-+8   130%    +9-+12
                If aObj.WieldReqVal = 0 Then
                    damageMod = 100
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 250 Then
                    damageMod = 110
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 270 Then
                    damageMod = 120
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 290 Then
                    damageMod = 130
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 315 Then
                    damageMod = 130
                    eDamage = 4
                ElseIf aObj.WieldReqVal = 335 Then
                    damageMod = 130
                    eDamage = 8
                ElseIf aObj.WieldReqVal >= 360 Then
                    damageMod = 130
                    eDamage = 12
                End If
            Case "Xbow"
                ' No      250     270     290     315         335     Bonus   360     Bonus
                ' 130%    140%    145%    155%    155% (+4)   155%    +5-+8   155%    +9-+12
                If aObj.WieldReqVal = 0 Then
                    damageMod = 130
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 250 Then
                    damageMod = 140
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 270 Then
                    damageMod = 145
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 290 Then
                    damageMod = 155
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 315 Then
                    damageMod = 155
                    eDamage = 4
                ElseIf aObj.WieldReqVal = 335 Then
                    damageMod = 155
                    eDamage = 8
                ElseIf aObj.WieldReqVal >= 360 Then
                    damageMod = 155
                    eDamage = 12
                End If
            Case "TW"
                ' No      250     270     290     315         335     Bonus   360     Bonus
                ' 120%    130%    140%    150%    150% (+4)   150%    +5-+8   150%    +9-+10
                If aObj.WieldReqVal = 0 Then
                    damageMod = 120
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 250 Then
                    damageMod = 130
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 270 Then
                    damageMod = 140
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 290 Then
                    damageMod = 150
                    eDamage = 0
                ElseIf aObj.WieldReqVal = 315 Then
                    damageMod = 150
                    eDamage = 4
                ElseIf aObj.WieldReqVal = 335 Then
                    damageMod = 150
                    eDamage = 8
                ElseIf aObj.WieldReqVal >= 360 Then
                    damageMod = 150
                    eDamage = 10
                End If
        End Select
    
        If damageMod > 0 Then
            'If aObj.TinkCount > 0 Then
            '    baseAL = baseAL + (20 * aObj.TinkCount)
            'End If
            
            'id = " (" & Round(((getRealArmorLevel(aObj) / baseAL) * 100), 1) & "%)"
            
            id = " (" & Round(((aObj.DamageModifier / damageMod) * 100), 1) & "%)"
        End If
    End If
    
    getUberRangedRating = id
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getUberRangedRating - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function


Public Function getUberMeleeRating(ByVal aObj As acObject) As String
On Error GoTo ErrorHandler

    Dim id As String
    Dim wieldReq, lowDamage, HighDamage As Integer

    wieldReq = 0
    lowDamage = 0
    HighDamage = 0
    
    MyDebug "getUberMeleeRating: " & aObj.Name & " WieldType: " & aObj.WieldReqType & " (" & aObj.WieldReqVal & ") :: " & g_ACConst.GetWeaponSkillName(aObj.WieldReqID)
    MyDebug "getUberMeleeRating: Dmg: " & aObj.GetLowDamage & "-" & aObj.HighDamage
    
    If (aObj.WieldReqVal > 0) And (aObj.WieldReqType = 2) Then
        Select Case g_ACConst.GetWeaponSkillName(aObj.WieldReqID)
            Case "Axe"
                '  0            250         300        325         350         370         400
                '11.4-19     16.8-28     19.2-32     21.6-36     22.8-38     25.2-42     27.6-46
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 11.4
                    HighDamage = 19
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 16.8
                    HighDamage = 28
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 19.2
                    HighDamage = 32
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 21.6
                    HighDamage = 36
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 22.8
                    HighDamage = 38
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 25.2
                    HighDamage = 42
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 27.6
                    HighDamage = 46
                End If
            Case "Dagger"
                '  0            250      300       325        350      370         400
                '8.4-12      9.1-13      9-15    12.6-18     14-20   16.8 -24    18.2-26
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 8.4
                    HighDamage = 12
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 9.1
                    HighDamage = 13
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 9
                    HighDamage = 15
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 12.6
                    HighDamage = 18
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 14
                    HighDamage = 20
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 16.8
                    HighDamage = 24
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 18.2
                    HighDamage = 26
                End If
            Case "Mace"
                '  0            250         300        325        350      370        400
                '13.5-18     19.5-26     22.5-30     25.5-34     27-36   28.5-38     31.5-42
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 13.5
                    HighDamage = 18
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 19.5
                    HighDamage = 26
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 22.5
                    HighDamage = 30
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 25.5
                    HighDamage = 34
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 27
                    HighDamage = 36
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 28.5
                    HighDamage = 38
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 31.5
                    HighDamage = 42
                End If
            Case "Spear"
                '  0            250         300        325         350         370        400
                '9.35-17     13.2-24     14.3-26     15.4-28     17.6-32     19.8-36     22-40
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 9.4
                    HighDamage = 17
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 13.2
                    HighDamage = 24
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 14.3
                    HighDamage = 26
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 15.4
                    HighDamage = 28
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 17.6
                    HighDamage = 32
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 19.8
                    HighDamage = 36
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 22
                    HighDamage = 40
                End If
            Case "Staff"
                '  0        250         300         325         350     370         400
                ' 9-12    9.75-14     11.25-15    14.75-18    15-20    18-24      19.5-26
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 9
                    HighDamage = 12
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 9.8
                    HighDamage = 14
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 11.3
                    HighDamage = 15
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 14.8
                    HighDamage = 18
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 15
                    HighDamage = 20
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 18
                    HighDamage = 24
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 19.5
                    HighDamage = 26
                End If
            Case "Sword"
                '  0            250         300        325         350         370       400
                ' 12-20        18-30       21-35      24-40       27-45       30-50     33-55
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 12
                    HighDamage = 20
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 18
                    HighDamage = 30
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 21
                    HighDamage = 35
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 24
                    HighDamage = 40
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 27
                    HighDamage = 45
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 30
                    HighDamage = 50
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 33
                    HighDamage = 55
                End If
            Case "UA"
                '  0       250         300        325       350      370       400
                ' 4-8     5.5-11      6.5-13      8-16     10-20    11-22     13-26
                If aObj.WieldReqVal = 0 Then
                    lowDamage = 4
                    HighDamage = 8
                ElseIf aObj.WieldReqVal = 250 Then
                    lowDamage = 5.5
                    HighDamage = 11
                ElseIf aObj.WieldReqVal = 300 Then
                    lowDamage = 6.5
                    HighDamage = 13
                ElseIf aObj.WieldReqVal = 325 Then
                    lowDamage = 8
                    HighDamage = 16
                ElseIf aObj.WieldReqVal = 350 Then
                    lowDamage = 10
                    HighDamage = 20
                ElseIf aObj.WieldReqVal = 370 Then
                    lowDamage = 11
                    HighDamage = 22
                ElseIf aObj.WieldReqVal >= 400 Then
                    lowDamage = 13
                    HighDamage = 26
                End If
        End Select
    
        If lowDamage > 0 Then
            'If aObj.TinkCount > 0 Then
            '    baseAL = baseAL + (20 * aObj.TinkCount)
            'End If
            
            'id = " (" & Round(((getRealArmorLevel(aObj) / baseAL) * 100), 1) & "%)"
            
            id = " (" & Round(((getRealMeleeDamage(aObj, False) / lowDamage) * 100), 1) & "%/" & Round(((getRealMeleeDamage(aObj, True) / HighDamage) * 100), 1) & "%)"
        End If
    End If
    
    getUberMeleeRating = id
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getUberMeleeRating - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Function getUberArmorRating(ByVal aObj As acObject) As String
On Error GoTo ErrorHandler

    Dim id As String
    Dim baseAL As Integer

    baseAL = 0
    
    MyDebug "getUberArmorRating: " & aObj.Name & " IT: " & aObj.itemType & " AL: " & aObj.ArmorLevel & " AT: " & aObj.ArmorType
    
    ' Max Armor levels
    If (aObj.itemType = ITEM_ARMOR) Or (aObj.ArmorLevel > 0) Then
        Select Case aObj.ArmorType
        Case ARMORTYPE_AMULI
            'Amuli , 254
            baseAL = 254
        Case ARMORTYPE_CELDON
            'Celdon , 274
            baseAL = 274
        Case ARMORTYPE_PLATEMAIL
            'Platemail , 264
            baseAL = 264
        Case ARMORTYPE_YOROI
            'Yoroi , 244
            baseAL = 244
        Case ARMORTYPE_KOUJIA
            'Koujia , 259
            baseAL = 259
        Case ARMORTYPE_LEATHER
            'Leather , 201
            baseAL = 201
        Case ARMORTYPE_SCALEMAIL
            'Scalemail , 239
            baseAL = 239
        Case ARMORTYPE_CHAINMAIL
            'Chainmail , 220
            baseAL = 220
        Case ARMORTYPE_COVENANT
            'Covenant , 499
            baseAL = 499
        Case ARMORTYPE_CHIRAN
            'Chiran , 264
            baseAL = 264
        Case ARMORTYPE_LORICA
            'Lorica , 265
            baseAL = 265
        Case ARMORTYPE_NARIYID
            'Nariyid , 274
            baseAL = 274
        Case ARMORTYPE_DIFORSA
            'Diforsa , 264
            baseAL = 264
        Case ARMORTYPE_ALDURESSA
            'Alduressa, 264
            baseAL = 264
        Case ARMORTYPE_TENASSA
            'Tenassa , 259
            baseAL = 259
        Case ARMORTYPE_BUCKLER
            'Buckler , 135
            baseAL = 135
        Case ARMORTYPE_KITESHIELD
            'Kite , 163
            baseAL = 163
        Case ARMORTYPE_ROUNDSHIELD
            'Round , 147
            baseAL = 147
        Case ARMORTYPE_TOWERSHIELD
            'Tower , 181
            baseAL = 181
        Case ARMORTYPE_COVENANTSHIELD
            'Covenant Shield , 399
            baseAL = 399
        Case Else
            'Misc armor's like helms, gloves, boots, etc
            '=========
            If InStr(1, LCase(aObj.Name), LCase("Armet")) Then
                'Armet , 315
                baseAL = 315
            ElseIf InStr(1, LCase(aObj.Name), LCase("Baigha")) Then
                'Baigha , 229
                baseAL = 229
            ElseIf InStr(1, LCase(aObj.Name), LCase("Bandana")) Then
                'Bandana , 292
                baseAL = 292
            ElseIf InStr(1, LCase(aObj.Name), LCase("Basinet")) Then
                'Basinet , 239
                baseAL = 239
            ElseIf InStr(1, LCase(aObj.Name), LCase("Beret")) Then
                'Beret , 292
                baseAL = 292
            ElseIf InStr(1, LCase(aObj.Name), LCase("Cap")) Then
                'Cap , 291
                baseAL = 291
            ElseIf InStr(1, LCase(aObj.Name), LCase("Circlet")) Then
                'Circlet , 239
                baseAL = 239
            ElseIf InStr(1, LCase(aObj.Name), LCase("Coif")) Then
                'Coif , 239
                baseAL = 239
            ElseIf InStr(1, LCase(aObj.Name), LCase("Coronet")) Then
                'Coronet , 260
                baseAL = 260
            ElseIf InStr(1, LCase(aObj.Name), LCase("Cowl")) Then
                'Cowl , 291
                baseAL = 291
            ElseIf InStr(1, LCase(aObj.Name), LCase("Crown")) Then
                'Crown , 304
                baseAL = 304
            ElseIf InStr(1, LCase(aObj.Name), LCase("Diadem")) Then
                'Diadem , 288
                baseAL = 288
            ElseIf InStr(1, LCase(aObj.Name), LCase("Fez")) Then
                'Fez , 277
                baseAL = 277
            ElseIf InStr(1, LCase(aObj.Name), LCase("Heaume")) Then
                'Heaume , 315
                baseAL = 315
            ElseIf InStr(1, LCase(aObj.Name), LCase("Horned Helm")) Then
                'Horned Helm, 252
                baseAL = 252
            ElseIf InStr(1, LCase(aObj.Name), LCase("Helmet")) Then
                'Helmet , 254
                baseAL = 254
            ElseIf InStr(1, LCase(aObj.Name), LCase("Kabuton")) Then
                'Kabuton , 254
                baseAL = 254
            ElseIf InStr(1, LCase(aObj.Name), LCase("Kasa")) Then
                'Kasa , 293
                baseAL = 293
            ElseIf InStr(1, LCase(aObj.Name), LCase("Qafiya")) Then
                'Qafiya , 293
                baseAL = 293
            ElseIf InStr(1, LCase(aObj.Name), LCase("Turban")) Then
                'Turban , 276
                baseAL = 276

            ElseIf InStr(1, LCase(aObj.Name), LCase("Kote")) Then
                'Kote , 264
                baseAL = 264

            ElseIf InStr(1, LCase(aObj.Name), LCase("Gauntlets")) Then
                'Gauntlets, 264 (platemail, if material not in piece name)
                baseAL = 264
            ElseIf InStr(1, LCase(aObj.Name), LCase("Gloves")) Then
                'Gloves, 243 (if soft - Leather, Satin, Velvet, etc)
                baseAL = 243

            ElseIf InStr(1, LCase(aObj.Name), LCase("Boots")) Then
                'Boots, 224 (if soft)
                baseAL = 224
            ElseIf InStr(1, LCase(aObj.Name), LCase("Loafers")) Then
                'Loafers , 243
                baseAL = 243
            ElseIf InStr(1, LCase(aObj.Name), LCase("Sandals")) Then
                'Sandals, 239 (if soft)
                baseAL = 239
            ElseIf InStr(1, LCase(aObj.Name), LCase("Shoes")) Then
                'Shoes , 243
                baseAL = 243
            ElseIf InStr(1, LCase(aObj.Name), LCase("Slippers")) Then
                'Slippers , 243
                baseAL = 243
            ElseIf InStr(1, LCase(aObj.Name), LCase("Sollerets")) Then
                'Sollerets, 264 (platemail)
                baseAL = 264
            End If
        End Select
        
        If baseAL > 0 Then
            If aObj.TinkCount > 0 Then
                baseAL = baseAL + (20 * aObj.TinkCount)
            End If
            id = " (" & Round(((getRealArmorLevel(aObj) / baseAL) * 100), 1) & "%)"
        End If
    End If
    
    
    getUberArmorRating = id
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLoot.getUberArmorRating - " & Err.Description & " - line: " & Erl
    Resume Fin
End Function

Public Sub exportInventory()
    Dim objItem As acObject
    
    PrintMessage "Starting ID of Inventory items..."
    
    'start a new ID queue
    Set m_IdQueue = New colObjects
    
    Call SetSilentID(True)
    For Each objItem In g_Objects.Items.Inv
        If Valid(objItem) Then
            Call m_IdQueue.Add(objItem.Guid, objItem.Name)
            'Call g_Service.IDObject(objItem.Guid)
            Call g_Hooks.IDQueueAdd(objItem.Guid)
        End If
    Next objItem
    
    m_idCount = m_IdQueue.Count + 1
    Call m_tmrExportInv.SetNextTime(1)

End Sub

Private Sub m_tmrExportInv_OnTimeOut()
    ' If the IDqueue count isn't moving, then we got stuck so bail
    If m_idCount = m_IdQueue.Count Then
        Call writeInventoryExport
        Exit Sub
    End If
    
    'Check to see if IDqueue is empty
    If m_IdQueue.Count > 0 Then
        m_idCount = m_IdQueue.Count
        Call m_tmrExportInv.SetNextTime(5)
        PrintMessage "Export: " & m_IdQueue.Count & " items left"
    Else
        Call writeInventoryExport
    End If
End Sub

Private Sub writeInventoryExport()
On Error GoTo ErrorHandler

    Dim oFS As New FileSystemObject
    Dim ts As TextStream
    
    Dim objItem As acObject
    Dim lngFileNr As Integer
    Dim sLine As String
    Dim sPath As String
    Dim sFileName As String
    
    sPath = g_Settings.GetDataFolder
    sFileName = g_ds.Player.Name & " Inventory.txt"
    
    MyDebug "clsLoot.exportInventory: sPath: " & sPath & "  sFileName: " & sFileName
    
    sFileName = sPath & "\" & sFileName
    
    ' Make a backup of each one
    If FileExists(sFileName) Then
        Call oFS.DeleteFile(sFileName, True)
    End If
    
    'MyDebug "Building " & sFileName
    
    Set ts = oFS.CreateTextFile(sFileName, False)
    
    If Not Valid(ts) Then
        ' Hmm, file is either aleady open or can't be written to, so bail!
         Exit Sub
    End If
    
    PrintMessage "Exporting Inventory to: " & sFileName
    
    ts.WriteLine ("====================================================================")
    ts.WriteLine ("======================= Inventory Export ===========================")
    ts.WriteLine ("====================================================================")
    
    'Armor's First
    ts.WriteBlankLines (2)
    ts.WriteLine ("-- Armors --")
    For Each objItem In g_Objects.Items.Inv
        If Not Valid(objItem) Then GoTo Fin
        If (objItem.itemType = ITEM_ARMOR) Then
            ts.WriteLine (" " & getIdString(objItem))
        End If
    Next objItem

    'Weapons
    ts.WriteBlankLines (2)
    ts.WriteLine ("-- Weapons --")
    For Each objItem In g_Objects.Items.Inv
        If Not Valid(objItem) Then GoTo Fin
        If (objItem.itemType = ITEM_MELEE_WEAPON) Or (objItem.itemType = ITEM_MISSILE_WEAPON) Then
            ts.WriteLine (" " & getIdString(objItem))
        End If
    Next objItem
    
    'Wands
    ts.WriteBlankLines (2)
    ts.WriteLine ("-- Wands --")
    For Each objItem In g_Objects.Items.Inv
        If Not Valid(objItem) Then GoTo Fin
        If (objItem.itemType = ITEM_WAND) Then
            ts.WriteLine (" " & getIdString(objItem))
        End If
    Next objItem
    
    'Underclothes
    ts.WriteBlankLines (2)
    ts.WriteLine ("-- Underclothes --")
    For Each objItem In g_Objects.Items.Inv
        If Not Valid(objItem) Then GoTo Fin
        If (objItem.itemType = ITEM_CLOTHING) Then
            ts.WriteLine (" " & getIdString(objItem))
        End If
    Next objItem
    
    'Jewelry
    ts.WriteBlankLines (2)
    ts.WriteLine ("-- Jewelry --")
    For Each objItem In g_Objects.Items.Inv
        If Not Valid(objItem) Then GoTo Fin
        If (objItem.itemType = ITEM_JEWELRY) Then
            ts.WriteLine (" " & getIdString(objItem))
        End If
    Next objItem
    
    'Rares
    ts.WriteBlankLines (2)
    ts.WriteLine ("-- Rares --")
    For Each objItem In g_Objects.Items.Inv
        If Not Valid(objItem) Then GoTo Fin
        If (objItem.IsRare) Then
            ts.WriteLine (" " & getIdString(objItem))
        End If
    Next objItem


    'All done, so close file
    ts.Close
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.exportInventory - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

'Local Debug
Private Sub locDebug(ByVal sDebugMsg As String, Optional ByVal bSilent As Boolean = False)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsLoot : " & GetStateString & "] " & sDebugMsg, bSilent)
    End If
End Sub


Public Sub DebugIgnoreCorpseList()
On Error GoTo ErrorHandler

    Dim Corpses() As Variant
    
1    Corpses = m_dicCorpseToIgnore.Keys
    Dim i As Long
    
    MyDebug "DebugIgnoreCorpseList - " & m_dicCorpseToIgnore.Count & " corpse(s) in list"
2    For i = LBound(Corpses) To UBound(Corpses)
3        MyDebug "Corpse #" & Corpses(i)
4        MyDebug "...." & g_Objects.FindObject(Corpses(i)).Name & " -> " & m_dicCorpseToIgnore(Corpses(i))  '(g_Core.Time - Val(m_dicCorpseToIgnore(Corpses(i))))
    Next i
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsLoot.DebugIgnoreCorpseList - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub
