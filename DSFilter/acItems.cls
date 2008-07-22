VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_colInv As colObjects
Private m_colWorld As colObjects
Private m_iNumTapers As Integer

'------------------ Not Functionnal ---------------------------------
Private m_iMainInventoryCount As Integer 'Number of items in the MAIN backpack
'------------------------ End ---------------------------------------

Private Const MAIN_PACK_CAPACITY = 102
Private Const MAIN_PACK_MAX_ITEMS = 96

'#####################################################################################
'#
'#                                      EVENTS
'#
'#####################################################################################

'Stack size of a stackable item has been adjusted
Public Event OnStackSizeChanged(ByVal objItem As acObject)

'We are equipping an item
Public Event OnEquipItem(ByVal objItem As acObject)

'An object belonging to the world has been inserted in our inventory
Public Event OnReceiveItem(ByVal objItem As acObject)

'An object belonging to our inventory has been handed out to someone else
Public Event OnGiveItem(ByVal objItem As acObject, ByVal objDest As acObject)

'An object we had equiped is being unequipied
Public Event OnUnequipItem(ByVal objItem As acObject)

'Setting the content of our backpacks or main inventory (objpack = objplayer) - collection only lists GUIDS / Item name
Public Event OnSetPackContent(ByVal objPack As acObject, ByVal colItems As colObjects)

'Receiving list of items in a chest/body etc - collection only lists GUIDS / Item name
Public Event OnOpenContainer(ByVal objContainer As acObject, ByVal colItems As colObjects)

'We are dropping an item on the ground
Public Event OnDropItemFromInventory(ByVal objItem As acObject)

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    Set m_colInv = New colObjects
    Set m_colWorld = New colObjects
    m_iNumTapers = 0
    m_iMainInventoryCount = 0
End Sub

Private Sub Class_Terminate()
    Set m_colInv = Nothing
    Set m_colWorld = Nothing
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get NumTapers() As Long
    NumTapers = m_iNumTapers
End Property

Public Property Get World() As colObjects
    Set World = m_colWorld
End Property

Public Property Get Inv() As colObjects
    Set Inv = m_colInv
End Property

Public Property Get MainInventoryCount() As Variant
    MainInventoryCount = m_iMainInventoryCount
End Property

'returns the percentage of encumbrement of the main backpack (100% = full, from the macro point of view)
Public Property Get MainInventoryPercent() As Integer
    MainInventoryPercent = CInt(Round(100 * (m_iMainInventoryCount / MAIN_PACK_MAX_ITEMS)))
End Property

Public Property Get BackpackFull() As Boolean
    BackpackFull = (m_iMainInventoryCount >= MAIN_PACK_MAX_ITEMS)
End Property

'#####################################################################################
'#
'#                               PRIVATE
'#
'#####################################################################################

Private Sub InventoryToWorld(ByVal lObjGUID As Long)
On Error GoTo ErrorHandler
    Dim objItem As acObject

    If m_colInv.Exists(lObjGUID) Then
        
        'Keep a reference to the item object
        Set objItem = m_colInv(lObjGUID)
        myDebug "[Inventory -> World] " & objItem.Name
    
        'remove item from inventory - we still own one reference to this object so it's not
        'deleted from memory
        Call m_colInv.Remove(lObjGUID)
        
        'Set it up to belong to the world
        objItem.Container = 0
        objItem.Wielder = 0
        
        'Put it into the World Collection
        Call AddWorldItemEx(objItem)
        
        'Update Inventory Count since the inventory has been modified
        Call UpdateInventoryCount
    Else
        myDebug "InventoryToWorld: obj #" & lObjGUID & " doesn't belong to Inventory"
    End If
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.InventoryToWorld - " & Err.Description
    Resume Fin
End Sub

Private Sub WorldToInventory(lObjGUID As Long, Optional lPackGUID As Long = 0)
On Error GoTo ErrorHandler
    Dim objItem As acObject

    If m_colWorld.Exists(lObjGUID) Then
        Dim bIsPack As Boolean
        
        'Keep a reference to this item
        Set objItem = m_colWorld(lObjGUID)
        myDebug "[World -> Inventory] " & objItem.Name
        
        'Remove item from world collection
        Call m_colWorld.Remove(lObjGUID)
        
        'TODO: add support for equip slots
        If lPackGUID = 0 Then lPackGUID = g_Objects.Player.GUID
        
        'Assign object to pack/player main inventory
        objItem.Container = lPackGUID
        
        'Move object to inventory
        Call AddInventoryItemEx(objItem)
        
        'Update Inventory Count since the inventory has been modified
        Call UpdateInventoryCount
    Else
        myDebug "WorldToInventory: obj #" & lObjGUID & " doesn't belong to m_colWorld"
    End If
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.InventoryToWorld - " & Err.Description
    Resume Fin
End Sub

'#####################################################################################
'#
'#                                PUBLIC
'#
'#####################################################################################

Public Function AddInventoryItem(ByVal itmName As String, ByVal itmGUID As Long) As acObject
On Error GoTo ErrorHandler
    
    myDebug "Inventory Items: added " & itmName
    Set AddInventoryItem = m_colInv.Add(itmGUID, itmName)
    Call UpdateInventoryCount

Fin:
    Exit Function
ErrorHandler:
    Set AddInventoryItem = Nothing
    myError "acItems.AddInventoryItem: " & Err.Description
    Resume Fin
End Function

Public Function AddInventoryItemEx(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    
    myDebug "Inventory Items: added " & objItem.Name
    AddInventoryItemEx = m_colInv.AddObject(objItem)
    Call UpdateInventoryCount

Fin:
    Exit Function
ErrorHandler:
    Set objItem = Nothing
    AddInventoryItemEx = False
    myError "acItems.AddInventoryItemEx: " & Err.Description
    Resume Fin
End Function

Public Function AddWorldItem(ByVal itmName As String, ByVal itmGUID As Long) As acObject
On Error GoTo ErrorHandler

    myDebug "AddWorldItem: m_colWorld Items: added " & itmName & " : " & itmGUID & "  (" & m_colWorld.Count & ")"
    Set AddWorldItem = m_colWorld.Add(itmGUID, itmName)

Fin:
    Exit Function
ErrorHandler:
    Set AddWorldItem = Nothing
    myError "acItems.AddWorldItem: " & Err.Description
    Resume Fin
End Function

Public Function AddWorldItemEx(objItem As acObject) As Boolean
On Error GoTo ErrorHandler
    
    myDebug "AddWorldItemEx: m_colWorld Items: added " & objItem.Name & " : " & objItem.GUID & "  (" & m_colWorld.Count & ")"
    AddWorldItemEx = m_colWorld.AddObject(objItem)

Fin:
    Exit Function
ErrorHandler:
    Set objItem = Nothing
    AddWorldItemEx = False
    myError "acItems.AddWorldItemEx: " & Err.Description
    Resume Fin
End Function

Public Function Exists(ByVal itmGUID As Long, Optional ByRef objItemOut As acObject) As Boolean
On Error GoTo ErrorHandler
    
    If m_colWorld.Exists(itmGUID, objItemOut) Then
        Exists = True
    ElseIf m_colInv.Exists(itmGUID, objItemOut) Then
        Exists = True
    Else
        Set objItemOut = Nothing
        Exists = False
    End If
    
Fin:
    Exit Function
ErrorHandler:
    Set objItemOut = Nothing
    Exists = False
    myError "acItems.Exists (" & itmGUID & ") - " & Err.Description
    Resume Fin
End Function

Public Function IsInMainInventory(objItem As acObject) As Boolean
On Error GoTo ErrorHandler

    IsInMainInventory = (objItem.Container = g_Objects.Player.GUID) _
                        And (Not objItem.IsPack) _
                        And (Not objItem.Equiped)
                        
Fin:
    Exit Function
ErrorHandler:
    IsInMainInventory = False
    myError "acItems.IsInMainInventory - " & Err.Description
    Resume Fin
End Function

Public Function Remove(ByVal itmGUID As Long) As Boolean
On Error GoTo ErrorHandler
    Dim bRemoved As Boolean
    
    If m_colWorld.Remove(itmGUID) Then
        bRemoved = True
    ElseIf m_colInv.Remove(itmGUID) Then
        bRemoved = True
        Call UpdateInventoryCount
    Else
        bRemoved = False
    End If
    
Fin:
    Remove = bRemoved
    Exit Function
ErrorHandler:
    bRemoved = False
    myError "acItems.Remove (" & itmGUID & ") - " & Err.Description
    Resume Fin
End Function

Public Function InInventory(ByVal lItemGUID As Long) As Boolean
On Error GoTo ErrorHandler
    
    InInventory = m_colInv.Exists(lItemGUID)

Fin:
    Exit Function
ErrorHandler:
    InInventory = False
    myError "acItems.InEnventory: " & Err.Description
    Resume Fin
End Function

Public Function InWorld(ByVal lItemGUID As Long) As Boolean
On Error GoTo ErrorHandler
    
    InWorld = m_colWorld.Exists(lItemGUID)

Fin:
    Exit Function
ErrorHandler:
    InWorld = False
    myError "acItems.InWorld: " & Err.Description
    Resume Fin
End Function

'Debug function
Public Sub DisplayInventory()
On Error GoTo ErrorHandler
    
    Dim tmpItem As acObject
    Dim Msg As String

    For Each tmpItem In m_colInv
        Msg = "     "
        If tmpItem.HasMinors Then Msg = Msg & " <<< MINOR >>> "
        If tmpItem.HasMajors Then Msg = Msg & " <<< MAJOR >>> "
        If tmpItem.IsRare Then Msg = Msg & " <<< RARE >>> "
        'Msg = Msg & tmpItem.Name & " [Type: " & tmpItem.GetTypeName & " (" & Hex(tmpItem.ItemType) & ") ] [Material: " & tmpItem.GetMaterialTypeName & "]"
        Msg = Msg & tmpItem.Name
        If tmpItem.Equiped Then Msg = Msg & " [Equiped]"
        myDebug Msg
    Next tmpItem

Fin:
    Set tmpItem = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.DisplayInventory: " & Err.Description
    Resume Fin
End Sub

'Debug function
Public Sub DisplayWorld()
    Dim tmpItem As acObject ', objOwner As acObject
    Dim Msg As String
    
    myDebug "Total: " & m_colWorld.Count & " world items"
    
    For Each tmpItem In m_colWorld
        Msg = "     "
        If tmpItem.HasMinors Then Msg = Msg & " <<< MINOR >>> "
        If tmpItem.HasMajors Then Msg = Msg & " <<< MAJOR >>> "
        If tmpItem.IsRare Then Msg = Msg & " <<< RARE >>> "
        Msg = Msg & tmpItem.Name
        
        'Set objOwner = FindOwner(tmpItem)
        'If (Not (objOwner Is Nothing)) Then
         '   Msg = Msg & " [Owner: " & objOwner.Name & "]"
        'End If
        
        Msg = Msg & " [Container: " & g_Objects.FindObject(tmpItem.Container).Name & "] [Owner: " & g_Objects.FindObject(tmpItem.Wielder).Name & "]"
        'Msg = Msg & " [" & tmpItem.GetTypeName & "] [Material: " & tmpItem.GetMaterialName & "]"
        
        myDebug Msg
    
    Next tmpItem
    
    Set tmpItem = Nothing
    'Set objOwner = Nothing

End Sub

Public Function Find(ByVal itmGUID As Long, Optional ByVal bReturnUnknown As Boolean = True) As acObject
Attribute Find.VB_UserMemId = 0
On Error GoTo ErrorHandler

    If m_colWorld.Exists(itmGUID) Then
        Set Find = m_colWorld(itmGUID)
    ElseIf m_colInv.Exists(itmGUID) Then
        Set Find = m_colInv(itmGUID)
    Else
        If bReturnUnknown Then
            g_Objects.Unknown.Name = "Unkown Item " & itmGUID
            Set Find = g_Objects.Unknown
        Else
            Set Find = Nothing
        End If
    End If
    
Fin:
    Exit Function
ErrorHandler:
    Find = Nothing
    myError "acItems.Find(" & itmGUID & ") - " & Err.Description
    Resume Fin
End Function

Public Function FindByName(ByVal sItemName As String, Optional ByVal bReturnUnknown As Boolean = True) As acObject
On Error GoTo ErrorHandler

    Dim objItem As acObject

    sItemName = Trim(LCase(sItemName))
    For Each objItem In m_colWorld
        If sItemName = LCase(objItem.Name) Then
            Set FindByName = objItem
            GoTo Fin
        End If
    Next objItem
    
    For Each objItem In m_colInv
        If sItemName = LCase(objItem.Name) Then
            Set FindByName = objItem
            GoTo Fin
        End If
    Next objItem
    
    'Item not found, return NULL
    If bReturnUnknown Then
        g_Objects.Unknown.Name = "Unknown Object " & sItemName
        Set FindByName = g_Objects.Unknown
    Else
        Set FindByName = Nothing
    End If
   
Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    FindByName = Nothing
    myError "acItems.FindByName(" & sItemName & ") - " & Err.Description
    Resume Fin
End Function

Public Function InvFindByName(ByVal sItemName As String, Optional ByVal bReturnUnknown As Boolean = True) As acObject
On Error GoTo ErrorHandler

    Dim objItem As acObject

    sItemName = Trim(LCase(sItemName))
    
    For Each objItem In m_colInv
        If sItemName = LCase(objItem.Name) Then
            Set InvFindByName = objItem
            GoTo Fin
        End If
    Next objItem
    
    'Item not found, return NULL
    If bReturnUnknown Then
        g_Objects.Unknown.Name = "Unknown Object " & sItemName
        Set InvFindByName = g_Objects.Unknown
    Else
        Set InvFindByName = Nothing
    End If

Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    InvFindByName = Nothing
    myError "acItems.InvFindByName(" & sItemName & ") - " & Err.Description
    Resume Fin
End Function

Public Function InvMatchByName(ByVal sName As String) As acObject
On Error GoTo ErrorHandler

    Dim objItem As acObject

    sName = Trim(LCase(sName))
    
    For Each objItem In m_colInv
        If InStr(1, LCase(objItem.Name), sName) > 0 Then
            Set InvMatchByName = objItem
            GoTo Fin
        End If
    Next objItem
    
    'not found
    Set InvMatchByName = Nothing
    
Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    InvMatchByName = Nothing
    myError "acItems.InvMatchByName(" & sName & ") - " & Err.Description
    Resume Fin
End Function

Public Function CountMainInventory() As Integer
On Error GoTo ErrorHandler

    Dim objItem As acObject
    
    CountMainInventory = 0
    For Each objItem In m_colInv
        If (objItem.Container = g_Objects.Player.GUID) _
        And (Not objItem.IsPack) _
        And (Not objItem.Equiped) Then
            CountMainInventory = CountMainInventory + 1
        End If
    Next objItem
    
Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    CountMainInventory = 0
    myError "acItems.CountMainInventory() - " & Err.Description
    Resume Fin
End Function

Public Sub UpdateInventoryCount()
    'update main pack count
    m_iMainInventoryCount = CountMainInventory
End Sub

Public Function InvCntByName(ByVal sItemName As String, Optional bExactMatch As Boolean = True) As Long
On Error GoTo ErrorHandler

    Dim objItem As acObject

    InvCntByName = 0
    sItemName = LCase(sItemName)
    
    For Each objItem In m_colInv
        If (Not bExactMatch And InStr(1, LCase(objItem.Name), sItemName) > 0) _
        Or (bExactMatch And (sItemName = LCase(objItem.Name))) Then
            InvCntByName = InvCntByName + objItem.StackCount
        End If
    Next objItem
    
Fin:
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    InvCntByName = 0
    myError "acItems.InvCntByName(" & sItemName & ") - " & Err.Description
    Resume Fin
End Function


'----------------------------------------------------------
Friend Sub NetEvent_OnAdjustStackSize(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim ObjectGUID As Long, NewCount As Long, NewValue As Long
    Dim objItem As acObject
    
    ObjectGUID = pMsg.Value("item")
    NewCount = pMsg.Value("count")
    NewValue = pMsg.Value("value")
    
    If Exists(ObjectGUID) Then
        Set objItem = Find(ObjectGUID)
        
        'MyDebug "OnAdjustStackSize: " & objItem.Name & " - New Count: " & lNewCount & " - New value: " & lNewValue
        objItem.Value = NewValue
        objItem.StackCount = NewCount

        RaiseEvent OnStackSizeChanged(objItem)
    End If
    
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnAdjustStackSize - " & Err.Description
    Resume Fin
End Sub

Friend Sub NetEvent_OnSetWielderContainer(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim ObjectGUID As Long, EquipType As Long, DestinationGUID As Long
    
    ObjectGUID = pMsg.Value("object")
    EquipType = pMsg.Value("equipType")
    DestinationGUID = pMsg.Value("container")
    
    If Exists(ObjectGUID) Then
    
        Dim objItem As acObject
        Dim objDest As acObject
        
        Set objItem = Find(ObjectGUID)
        Set objDest = g_Objects.FindObject(DestinationGUID)
        
        If EquipType = 2 Then   'Set container
            'MyDebug "OnSetWielderContainer : " & objItem.Name & " - Setting Container to " & objDest.Name
            
            'check to see if someone's giving us an item
            If objDest.GUID = g_Objects.Player.GUID Then
                If Not InInventory(ObjectGUID) Then
                    myDebug "OnSetWielderContainer: someone's giving us " & objItem.Name
                    Call WorldToInventory(ObjectGUID)
                    Set objItem = Find(ObjectGUID)
                    objItem.Wielder = 0
                End If
            Else
                objItem.Container = DestinationGUID
            End If
            
        ElseIf EquipType = 3 Then 'Set Wielder
                        
            'WARNING: if this item's wielder/owner is a monster, and the new wielder is set to world
            'it means that the monster is switching weapon, and putting the current one in his backpack
            'For the players, the client would be receiving 2 SetWielderContainer messages (one to
            'set the new wielder, and another one to set the container) but it's different for the monsters
            'Indeed, monsters will never drop their weapons on the flour by themselves, so only
            '1 SetWielderContainer message is required, to tell the client it's removing his current
            'Weapon. Thus, we have to set the current item's container to the guid of the monster
            'so that it remains linked to him.
            
            If objItem.IsWielded And (DestinationGUID = 0) Then
                Dim objOwner As acObject
                Set objOwner = g_Objects.FindObject(objItem.Wielder)
                If g_Objects.Monsters.Exists(objOwner.GUID) Then
                    myDebug "OnSetWielderContainer: monster " & objOwner.Name & " unequipped " & objItem.Name & " - Updating container info."
                    objItem.Container = objOwner.GUID
                End If
                Set objOwner = Nothing
            End If
            
            objItem.Wielder = DestinationGUID
            
            'MyDebug "OnSetWielderContainer : " & objItem.Name & " - Setting Wielder to " & g_Objects.FindObject(DestinationGUID).Name
        
        End If
        
        Set objDest = Nothing
        Set objItem = Nothing
    
    Else
        Dim sType As String
        If EquipType = 2 Then
            sType = "Container"
        Else
            sType = "Wielder"
        End If
        myDebug "OnSetWielderContainer : " & g_Objects.FindObject(ObjectGUID).Name & " doesnt exist in m_colWorld/Inventory - Tryed to set " & sType & " to " & g_Objects.FindObject(DestinationGUID).Name
    End If
    
    'm_colInv count changed
    Call UpdateInventoryCount
    
Fin:
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnSetWielderContainer - " & Err.Description
    Resume Fin
End Sub

Friend Sub NetEvent_OnWieldObject(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim ObjectGUID As Long, WielderGUID As Long
    ObjectGUID = pMsg.Value("object")
    WielderGUID = pMsg.Value("owner")
    
    If Exists(ObjectGUID) Then
        Dim objItem As acObject
        Set objItem = Find(ObjectGUID, False)
        
        If Valid(objItem) Then
            myDebug "OnWieldObject: " & objItem.Name
            objItem.Wielder = WielderGUID
            
            If objItem.Wielder = g_Objects.Player.GUID Then
                objItem.Equiped = True
                RaiseEvent OnEquipItem(objItem)
            Else    'object equiped by someone else
                objItem.Equiped = False
            End If
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnWieldObject - " & Err.Description
    Resume Fin
End Sub

Friend Sub NetEvent_OnMoveObjectIntoInventory(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim ObjectGUID As Long
    ObjectGUID = pMsg.Value("object")
    
    If Exists(ObjectGUID) Then
        Dim objItem As acObject
        Set objItem = Find(ObjectGUID, False)
        
        If Valid(objItem) Then
            'MyDebug "MoveObjectIntoInventory: " & objItem.Name
        
            If Not InInventory(ObjectGUID) Then
                If objItem.IsWielded Then
                    'if item was previously wielded by someone, it means the object has been put back into this
                    'person's inventory -> we have to swap Owner and Container
                    'MyDebug "MoveObjectIntoInventory: " & objItem.Name & " - Moved item back into its wielder's inventory"
                    objItem.Container = objItem.Wielder
                ElseIf objItem.IsOnGround Then
                    'MyDebug "MoveObjectIntoInventory: " & objItem.Name & " picked up!"
                    'we don't know who picked it up, we only know that this item is no longer wielded
                    'and no longer on ground
                    objItem.Container = -1
                End If
            End If
            
            'moved in inventory, so no longer wielded
            objItem.Wielder = 0
        
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnMoveObjectIntoInventory - " & Err.Description
    Resume Fin
End Sub

Friend Sub NetEvent_OnInsertInventoryItem(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim ObjectGUID As Long, DestinationGUID As Long, Slot As Long, SlotKind As Long
    Dim objItem As acObject, objDest As acObject
   
    ObjectGUID = pMsg.Value("item")
    DestinationGUID = pMsg.Value("container")
    Slot = pMsg.Value("slot")
    SlotKind = pMsg.Value("type")
    
    Set objDest = g_Objects.FindObject(DestinationGUID)
    Set objItem = Find(ObjectGUID, False)
        
    If Valid(objItem) Then
        myDebug "OnInsertInventoryItem: " & objItem.Name & " -- Dest : " & objDest.Name & " - Slot: " & Slot & " [type:" & SlotKind & "]"
        
        'If object's coming from outside (world),
        'move it to our inventory collection and remove it from world
        If InWorld(ObjectGUID) Then
        
            'first check if the destination is us, or something belonging to us
            '(ie: a pack)
            If (objDest.GUID = g_Objects.Player.GUID) Or (objDest.Container = g_Objects.Player.GUID) Then
                myDebug "OnInsertInventoryItem: receiving/picked up " & objItem.Name
                
                'Move item from m_colWorld collection to Inventory collection
                Call WorldToInventory(ObjectGUID, DestinationGUID)
                
                'Fire the event
                RaiseEvent OnReceiveItem(objItem)
                
            Else    'destination is another object in the world
                myDebug "OnInsertInventoryItem: " & objDest.Name & " receives " & objItem.Name
            End If
            
        Else 'item is already in our inventory
        
            'if destination is not main pack (player.guid), and
            'doesnt belong to us (ie: not a pack), then it means
            'this item is given to someone else around us
            If (objDest.GUID <> g_Objects.Player.GUID) And (objDest.Container <> g_Objects.Player.GUID) Then
                'MyDebug "OnInsertInventoryItem: giving " & objItem.Name & " to " & objDest.Name
                Call InventoryToWorld(ObjectGUID)
                
                'Fire the event
                RaiseEvent OnGiveItem(objItem, objDest)
            End If
            
        End If 'Item in world
        
        'FIXME: Equip flag depends of destination slot
        If objItem.Equiped Then
            'MyDebug "OnInsertInventoryItem: unequipping " & objItem.Name
            objItem.Equiped = False
            RaiseEvent OnUnequipItem(objItem)
        End If
        
        'Set new container
        objItem.Container = DestinationGUID

    Else
        myDebug "OnInsertInventoryItem(" & g_Objects.FindObject(ObjectGUID).Name & " -> " & objDest.Name & ") - Item not in m_colWorld/Inventory"
    End If
        
Fin:
    Set objItem = Nothing
    Set objDest = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnInsertInventoryItem - " & Err.Description
    Resume Fin
End Sub


Friend Sub NetEvent_OnWearItem(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler
    
    Dim ObjectGUID As Long, CoverageSlot As Long
    Dim objItem As acObject
    
    ObjectGUID = pMsg.Value("item")
    CoverageSlot = pMsg.Value("slot")
    
    Set objItem = Find(ObjectGUID, False)
    If Valid(objItem) Then
        myDebug "OnWearItem: equipping " & objItem.Name
        objItem.Equiped = True
        
        'Fire the event
        RaiseEvent OnEquipItem(objItem)
    Else
        myDebug "OnWearItem - Unknown item"
    End If
    
    Call UpdateInventoryCount
         
Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnWearItem - " & Err.Description
    Resume Fin
End Sub

Friend Sub NetEvent_OnSetPackContents(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim PackGUID As Long, ItemCount As Long
    Dim objContainer As acObject
    Dim i As Integer
    Dim lObjId As Long
    Dim colItems As New colObjects
    
    PackGUID = pMsg.Value("container")
    ItemCount = pMsg.Value("itemCount")
    
    Set objContainer = g_Objects.FindObject(PackGUID)
    myDebug "OnSetPackContent - Container : " & objContainer.Name & " -> " & ItemCount & " items"

    'List of items in container
    'We only know the items GUIDs right now, the OnCreateObject message will come right after
    For i = 0 To ItemCount - 1
        With pMsg.Struct("items").Struct(i)
            lObjId = .Value("item")
            Call colItems.Add(lObjId, "Item [" & lObjId & "] #" & i & " on " & objContainer.Name)
        End With
    Next i
    
    'If this is one of our packs
    If (objContainer.GUID = g_Objects.Player.GUID) Or (objContainer.Container = g_Objects.Player.GUID) Then
        RaiseEvent OnSetPackContent(objContainer, colItems)
    Else
        RaiseEvent OnOpenContainer(objContainer, colItems)
    End If
   
Fin:
    Set objContainer = Nothing
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnSetPackContents - " & Err.Description
    Resume Fin
End Sub

Friend Sub NetEvent_OnDropFromInventory(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim ObjectGUID As Long
    ObjectGUID = pMsg.Value("item")
    
    'myDebug "OnDropItem : " & g_Objects.FindObject(ObjectGUID).Name
    Dim objItem As acObject
    Set objItem = Find(ObjectGUID, False)
    
    If Valid(objItem) Then
        RaiseEvent OnDropItemFromInventory(objItem)
    End If
    
    Call InventoryToWorld(ObjectGUID)
   
Fin:
    Exit Sub
ErrorHandler:
    myError "acItems.NetEvent_OnDropFromInventory - " & Err.Description
    Resume Fin
End Sub

Friend Function CreateItem(ByVal ObjectName As String, ByVal ObjectGUID As Long, ByVal ContainerGUID As Long, ByVal WielderGUID As Long, Optional ByVal bIsPack As Boolean = False) As acObject
On Error GoTo ErrorHandler
    
    Dim objItem As acObject
    
    myDebug "acItems.CreateItem - " & ObjectName & " [" & ObjectGUID & "]"
    
    'If object belongs to us
    If (ContainerGUID = g_Objects.Player.GUID) Or (WielderGUID = g_Objects.Player.GUID) Then
        Set objItem = AddInventoryItem(ObjectName, ObjectGUID)
    Else
        'The current object's Container can be a Pack in our inventory. Check to see if it's the case
        If InInventory(ContainerGUID) Then 'object belongs to one our packs, add it to inventory
            Set objItem = AddInventoryItem(ObjectName, ObjectGUID)
        Else    'world object
            Set objItem = AddWorldItem(ObjectName, ObjectGUID)
        End If
    End If
    
    objItem.ObjectType = TYPE_ITEM
    objItem.Wielder = WielderGUID
    objItem.Container = ContainerGUID
    
    'myDebug "CreateItem: " & objItem.Name & " [" & objItem.GUID & "]"
    
Fin:
    Set CreateItem = objItem
    Set objItem = Nothing
    Exit Function
ErrorHandler:
    Set objItem = Nothing
    myError "acItems.CreateItem - " & Err.Description
    Resume Fin
End Function
