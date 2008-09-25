VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_colPlayers As colObjects      'the other players around us
Private m_colMonsters As colObjects     'collection of monsters we are currently aware of
Private m_colJunk As colObjects         'collection for other junk objects
Private m_colVendorInv As colObjects    'collection of this vendors inventory

Private m_Items As acItems         'world/inventory items
Private m_Equipment As acEquipment 'collection items worn

'Special entities
Private m_objPlayer As acObject     'self/local player (macro) - points to a colPlayers item
Private m_objUnknown As acObject    'unknown object
Private m_objWorld As acObject      'world entity
Private m_objSelected As acObject   'selected object
Private m_objVendor As acObject     'currently selected vendor

Private WithEvents m_Fellowship As acFellowship 'our fellowship if any
Attribute m_Fellowship.VB_VarHelpID = -1

Public Event OnRemoveObject(ByVal obj As acObject)
Public Event OnCreateObject(ByVal obj As acObject)
Public Event OnObjectMoved(ByVal obj As acObject)
Public Event OnIdentifyObject(ByVal obj As acObject)
Public Event OnSelectObject(ByVal obj As acObject)
Public Event OnApproachVendor(ByVal obj As acObject, ByVal vendorInv As colObjects)

'Externally Triggered Events
Public Event OnObjectVulned(ByVal obj As acObject, ByVal iVuln As Integer)
Public Event OnObjectImpOrBludgeoned(ByVal obj As acObject)
Public Event OnObjectYielded(ByVal obj As acObject)
Public Event OnObjectDeath(ByVal obj As acObject)
Public Event OnVulnExpired(ByVal obj As acObject)

Public Event OnMinorDetected(ByVal obj As acObject, ByVal sSpellName As String)
Public Event OnMajorDetected(ByVal obj As acObject, ByVal sSpellName As String)
Public Event OnEpicDetected(ByVal obj As acObject, ByVal sSpellName As String)
Public Event OnRareDetected(ByVal obj As acObject)


'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    Set m_objUnknown = New acObject
        m_objUnknown.Name = "UnknownObject"
        m_objUnknown.GUID = -1
        
    Set m_objWorld = New acObject
        m_objWorld.Name = "World"
        m_objWorld.GUID = 0
    
    Set m_objPlayer = New acObject
        m_objPlayer.Name = "LocalPlayer"        'temporary
        m_objPlayer.GUID = 0
        
    Set m_objVendor = New acObject
        m_objVendor.Name = "None"
        m_objVendor.GUID = 0
        
    Set m_colVendorInv = New colObjects
        m_colVendorInv.Description = "Vendor Inventory"
    
    Set m_colJunk = New colObjects
        m_colJunk.Description = "Junk Objects"
        
    Set m_colMonsters = New colObjects
        m_colMonsters.Description = "Monsters"
        
    Set m_colPlayers = New colObjects
        m_colPlayers.Description = "Players"
        
    Set m_Fellowship = New acFellowship
    Set m_Items = New acItems
    Set m_Equipment = New acEquipment
    
    Set m_objSelected = Nothing

Fin:
    Exit Sub
ErrorHandler:
    myError "GameEntities.Class_Inititalize"
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set m_objUnknown = Nothing
    Set m_objWorld = Nothing
    Set m_objPlayer = Nothing
    Set m_objSelected = Nothing
    Set m_objVendor = Nothing
    Set m_colVendorInv = Nothing
    Set m_colMonsters = Nothing
    Set m_colPlayers = Nothing
    Set m_colJunk = Nothing
    Set m_Fellowship = Nothing
    Set m_Items = Nothing
    Set m_Equipment = Nothing
End Sub

'#####################################################################################
'#
'#                                  PROPERTIES
'#
'#####################################################################################

Public Property Get Players() As colObjects
    Set Players = m_colPlayers
End Property

Public Property Get Monsters() As colObjects
    Set Monsters = m_colMonsters
End Property

Public Property Get Player() As acObject
    Set Player = m_objPlayer
End Property

Public Property Get Unknown() As acObject
    Set Unknown = m_objUnknown
End Property

Public Property Get World() As acObject
    Set World = m_objWorld
End Property

Public Property Get Vendor() As acObject
    Set Vendor = m_objVendor
End Property

Public Property Get vendorInv() As colObjects
    Set vendorInv = m_colVendorInv
End Property

Public Property Get Junk() As colObjects
    Set Junk = m_colJunk
End Property

Public Property Get Selected() As acObject
    If g_Filter.ds_ACHooks.CurrentSelection <> 0 Then
        Set Selected = m_objSelected
    Else
        Set Selected = Nothing
    End If
End Property

Public Property Get Fellowship() As acFellowship
    Set Fellowship = m_Fellowship
End Property

Public Property Get Items() As acItems
    Set Items = m_Items
End Property

Public Property Get Equipment() As acEquipment
    Set Equipment = m_Equipment
End Property

'#####################################################################################
'#
'#                                  Event Triggers
'#
'#####################################################################################

Friend Sub FireObjectVulned(ByVal obj As acObject, ByVal iVuln As Integer)
    If Valid(obj) Then RaiseEvent OnObjectVulned(obj, iVuln)
End Sub

Friend Sub FireObjectImpOrBludgeoned(ByVal obj As acObject)
    If Valid(obj) Then RaiseEvent OnObjectImpOrBludgeoned(obj)
End Sub

Friend Sub FireOnObjectYielded(ByVal obj As acObject)
    If Valid(obj) Then RaiseEvent OnObjectYielded(obj)
End Sub

Friend Sub FireOnObjectDeath(ByVal obj As acObject)
    If Valid(obj) Then RaiseEvent OnObjectDeath(obj)
End Sub

'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

Private Function CreatePlayer(ByVal sObjName As String, ByVal lObjGUID As Long) As acObject
On Error GoTo ErrorHandler

    Dim objPlayer As acObject
    
    myDebug "GameObjects.CreatePlayer : " & sObjName
    If lObjGUID <> m_objPlayer.GUID Then
        myDebug "[CreatePlayer] Creating Player : " & sObjName & " [" & lObjGUID & "]"
        Set objPlayer = m_colPlayers.Add(lObjGUID, sObjName)
    Else
        myDebug "[CreatePlayer] Creating Local Player : " & sObjName & " [" & lObjGUID & "]"
        Set objPlayer = m_objPlayer
        objPlayer.Name = sObjName
    End If
    
    objPlayer.ObjectType = TYPE_PLAYER
    
Fin:
    Set CreatePlayer = objPlayer
    Set objPlayer = Nothing
    Exit Function
ErrorHandler:
    Set objPlayer = Nothing
    myError "GameEntities.CreatePlayer - " & Err.Description
    Resume Fin
End Function

Private Function CreateMonster(ByVal sObjName As String, ByVal lObjGUID As Long) As acObject
On Error GoTo ErrorHandler

    Dim objMonster As acObject
    Set objMonster = m_colMonsters.Add(lObjGUID, sObjName)
    objMonster.ObjectType = TYPE_MONSTER
   
Fin:
    Set CreateMonster = objMonster
    Set objMonster = Nothing
    Exit Function
ErrorHandler:
    Set objMonster = Nothing
    myError "GameEntities.CreateMonster - " & Err.Description
    Resume Fin
End Function

Private Function CreateJunk(ByVal sObjName As String, ByVal lObjGUID As Long) As acObject
On Error GoTo ErrorHandler

    Dim objJunk As acObject
    Set objJunk = m_colJunk.Add(lObjGUID, sObjName)
    objJunk.ObjectType = TYPE_JUNK
   
Fin:
    Set CreateJunk = objJunk
    Exit Function
ErrorHandler:
    Set objJunk = Nothing
    myError "GameEntities.CreateJunk - " & Err.Description
    Resume Fin
End Function

'#####################################################################################
'#
'#                                  PUBLIC
'#
'#####################################################################################

Public Function IsSelf(obj As acObject) As Boolean
    IsSelf = False
    If Valid(obj) Then IsSelf = (obj Is m_objPlayer) Or (obj.GUID = m_objPlayer.GUID)
End Function

Public Function IsUnknown(obj As acObject) As Boolean
    IsUnknown = False
    If Valid(obj) Then IsUnknown = (obj Is m_objUnknown)
End Function

Public Function IsWorld(obj As acObject) As Boolean
    IsWorld = False
    If Valid(obj) Then IsWorld = (obj Is m_objWorld) Or (obj.GUID = m_objWorld.GUID)
End Function

Public Function IsInPlayersCol(obj As acObject) As Boolean
    IsInPlayersCol = False
    If Valid(obj) Then IsInPlayersCol = m_colPlayers.Exists(obj.GUID)
End Function

Public Function IsInMonstersCol(obj As acObject) As Boolean
    IsInMonstersCol = False
    If Valid(obj) Then IsInMonstersCol = m_colMonsters.Exists(obj.GUID)
End Function

Private Sub RemoveWieldedItems(ByVal objEntity As acObject)
On Error GoTo ErrorHandler
    
    If Not Valid(objEntity) Then
        myError "GameObjects.RemoveWieldedItems: invalid objEntity - ignoring"
        GoTo Fin
    End If
    
    Dim objItem As acObject
    For Each objItem In m_Items.World
        If objItem.Wielder = objEntity.GUID Then
            myDebug "GameObjects.RemoveWieldedItems: " & objItem.Name & " (from " & objEntity.Name & ")"
            If m_Items.Remove(objItem.GUID) Then
                RaiseEvent OnRemoveObject(objItem)
            Else
                myError "GameObjects.RemoveWieldedItems - Failed to remove object " & objItem.Name & " from collection"
            End If
        End If
    Next objItem

Fin:
    Set objItem = Nothing
    Exit Sub
ErrorHandler:
    myError "GameObjects.RemoveWieldedItems - " & Err.Description
    Resume Fin
End Sub

Friend Sub RemoveObject(ByVal lObjGUID As Long, Optional ByVal aSource As String)
On Error GoTo ErrorHandler

    'Never remove self
    If lObjGUID = m_objPlayer.GUID Then
        myError "GameObjects.RemoveObject : tryed to remove local player object [id: " & lObjGUID & "]. Ignoring."
        GoTo Fin
    End If
    
    Dim objEntity As acObject
    Dim bRemoved As Boolean

    If Not Valid(g_Objects.FindObject(lObjGUID)) Then Exit Sub
    
    myDebug "RemoveObject: (" & aSource & "): " & g_Objects.FindObject(lObjGUID).Name & " : " & lObjGUID
    
    'Try to find the object
    'If we have a match, keep a reference to this object, and remove it from the appropriate collection
    'Since we keep a reference to the object, it won't be removed from memory yet so we can still use it
    'in this function
    If m_colMonsters.Exists(lObjGUID, objEntity) Then
        bRemoved = m_colMonsters.Remove(lObjGUID)
        myDebug "RemoveObject: removed from colMonsters: " & lObjGUID
    ElseIf m_colPlayers.Exists(lObjGUID, objEntity) Then
        bRemoved = m_colPlayers.Remove(lObjGUID)
        myDebug "RemoveObject: removed from colPlayers: " & lObjGUID
    ElseIf m_Items.Exists(lObjGUID, objEntity) Then
        bRemoved = m_Items.Remove(lObjGUID)
        myDebug "RemoveObject: removed from colItems: " & lObjGUID
    ElseIf m_colJunk.Exists(lObjGUID, objEntity) Then
        bRemoved = m_colJunk.Remove(lObjGUID)
        myDebug "RemoveObject: removed from colJunk: " & lObjGUID
    End If
    
    'If we found something
    If Valid(objEntity) Then
        'Remove wielded objects from player or monsters
        If (objEntity.ObjectType = TYPE_PLAYER) Or (objEntity.ObjectType = TYPE_MONSTER) Then
            Call RemoveWieldedItems(objEntity)
        End If
        If bRemoved Then
            RaiseEvent OnRemoveObject(objEntity)
        Else
            myError "GameObjects.RemoveObject - Failed to remove object " & objEntity.Name & " from collection"
        End If
    End If
    
Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "GameObjects.RemoveObject - " & Err.Description
    Resume Fin
End Sub

Public Function Exists(ByVal lObjGUID As Long, Optional ByRef objOut As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If lObjGUID = 0 Then
        Set objOut = m_objWorld
        bRet = True
    ElseIf lObjGUID = m_objPlayer.GUID Then
        Set objOut = m_objPlayer
        bRet = True
    ElseIf m_colMonsters.Exists(lObjGUID, objOut) Then
        bRet = True
    ElseIf m_colPlayers.Exists(lObjGUID, objOut) Then
        bRet = True
    ElseIf m_Items.Exists(lObjGUID, objOut) Then
        bRet = True
    ElseIf m_colJunk.Exists(lObjGUID, objOut) Then
        bRet = True
    Else    ' not found
        bRet = False
        Set objOut = Nothing
    End If
    
Fin:
    Exists = bRet
    Exit Function
ErrorHandler:
    Set objOut = Nothing
    bRet = False
    myError "GameObjects.Exists - " & Err.Description
    Resume Fin
End Function

Public Function FindObject(ByVal lObjGUID As Long, Optional ByVal bReturnUnknown As Boolean = True) As acObject
On Error GoTo ErrorHandler

    Dim objEntity As acObject
    
    If Exists(lObjGUID, objEntity) Then
        Set FindObject = objEntity
    Else
        If bReturnUnknown Then
            m_objUnknown.GUID = -1
            m_objUnknown.Name = "Unknown Object " & lObjGUID
            Set FindObject = m_objUnknown
        Else
            Set FindObject = Nothing
        End If
    End If

Fin:
    Set objEntity = Nothing
    Exit Function
ErrorHandler:
    Set FindObject = m_objUnknown
    myError "GameObjects.FindObject - " & Err.Description
    Resume Fin
End Function

Public Function GetModObject(ByVal lObjGUID As Long, Optional ByRef objOut As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If lObjGUID = 0 Then
        Set objOut = m_objWorld
        bRet = True
    ElseIf lObjGUID = m_objPlayer.GUID Then
        Set objOut = m_objPlayer
        bRet = True
    ElseIf m_colMonsters.Exists(lObjGUID, objOut) Then
        bRet = True
    ElseIf m_colPlayers.Exists(lObjGUID, objOut) Then
        bRet = True
    ElseIf m_Items.Exists(lObjGUID, objOut) Then
        bRet = True
    ElseIf m_colJunk.Exists(lObjGUID, objOut) Then
        bRet = True
    Else    ' not found
        bRet = False
        Set objOut = Nothing
    End If
    
Fin:
    GetModObject = bRet
    Exit Function
ErrorHandler:
    Set objOut = Nothing
    bRet = False
    myError "GameObjects.GetModObject - " & Err.Description
    Resume Fin
End Function

Public Function FindPlayer(ByVal lPlayerGUID As Long) As acObject
On Error GoTo ErrorHandler

    If lPlayerGUID = m_objPlayer.GUID Then
        Set FindPlayer = m_objPlayer
        
    ElseIf m_colPlayers.Exists(lPlayerGUID) Then
        Set FindPlayer = m_colPlayers(lPlayerGUID)
        
    Else
        Set FindPlayer = Nothing
    End If

Fin:
    Exit Function
ErrorHandler:
    Set FindPlayer = Nothing
    myError "GameObjects.FindPlayer - " & Err.Description
    Resume Fin
End Function

'Returns object if found, or Nothing if not found
Public Function FindPlayerByName(ByVal sPlayerName As String) As acObject
On Error GoTo ErrorHandler

    Dim objPlayer As acObject

    If SameText(sPlayerName, m_objPlayer.Name) Then
        Set objPlayer = m_objPlayer
    Else
        For Each objPlayer In m_colPlayers
            If SameText(objPlayer.Name, sPlayerName) Then
                GoTo Fin    'exit for
            End If
        Next objPlayer
        
        'not found
        Set objPlayer = Nothing
    End If
    
Fin:
    Set FindPlayerByName = objPlayer
    Set objPlayer = Nothing
    Exit Function
ErrorHandler:
    Set objPlayer = Nothing
    myError "FindPlayerByName"
    Resume Fin
End Function

'#####################################################################################
'#
'#                     Object Events (triggered by NetEcho)
'#
'#####################################################################################

Friend Sub NetEvent_OnCreateObject(ByVal pMsg As DecalNet.IMessage2, ByVal isUpdate As Boolean)
On Error GoTo ErrorHandler

    Dim objEntity As acObject
    Dim lObjGUID As Long
    Dim sObjName As String
    Dim gameFlags1 As Long, physFlag As Long, catFlag As Long, behaveFlag As Long
    Dim xGame As DecalNet.IMessageMember
    Dim xPhysics As DecalNet.IMessageMember
    Dim bAlreadyExist As Boolean
    
    lObjGUID = pMsg.Value("object")
    
    Set xGame = pMsg.Struct("game")
      
    sObjName = CStr(xGame.Value("name"))
    gameFlags1 = xGame.Value("flags1")
    catFlag = xGame.Value("category")
    behaveFlag = xGame.Value("behavior")
    
    'myDebug "NetEvent_OnCreateObject: " & lObjGUID & " : " & sObjName

' Don't care about duplicates, as colObjects.addObject now just removes the old one and adds the new one
    '
    'Make sure we're not trying to add an already existing object that hasn't been deleted (happens sometimes)
    If isUpdate Then
        Set objEntity = FindObject(lObjGUID, False)
        If Not Valid(objEntity) Then
            myDebug "OnCreate: isUpdate = TRUE, but can not find existing object"
            Set objEntity = New acObject
        Else
            'Hmm, do we need to do anything here?
        End If
    Else
        ' See if we need to remove it from the other collections first
        If Valid(g_Objects.FindObject(lObjGUID)) Then
        '    Set objEntity = FindObject(lObjGUID, False)
        '    'If m_colMonsters.Exists(lObjGUID, objEntity) Then
        '    '    Call m_colMonsters.Remove(lObjGUID)
        '    '    myError "OnCreateObject: DUPLICATE GUID in colMonsters: " & lObjGUID
        '    'ElseIf m_colPlayers.Exists(lObjGUID, objEntity) Then
        '    '    Call m_colPlayers.Remove(lObjGUID)
        '    '    myDebug "OnCreateObject: DUPLICATE GUID in colPlayers: " & lObjGUID
            If m_Items.Exists(lObjGUID) Then
                Call m_Items.Remove(lObjGUID)
                myDebug "OnCreateObject: DUPLICATE GUID in colItems: " & lObjGUID
            ElseIf m_colJunk.Exists(lObjGUID) Then
                Call m_colJunk.Remove(lObjGUID)
                myDebug "OnCreateObject: DULICATE GUID in colJunk: " & lObjGUID
            End If
        End If

        ' Ok, create a new one
        Set objEntity = New acObject
    
    End If
    
    objEntity.GUID = lObjGUID
    objEntity.Name = sObjName
    objEntity.Icon = xGame.Value("icon")
    objEntity.gameDataType = xGame.Value("type")
    
    If Not Valid(objEntity) Then
        myError "NetEvent_OnCreateObject: not VALID objEntity: " & sObjName & " :guid: " & lObjGUID
    End If
    
    myDebug "NetEvent_OnCreateObject: " & objEntity.Name & " [" & objEntity.GUID & "] - gameFlags1 =" & Hex(gameFlags1)
    
    'Location
    Set xPhysics = pMsg.Struct("physics")
    physFlag = xPhysics.Value("flags")
    
    If (physFlag And &H20) Then objEntity.Wielder = xPhysics.Value("equipper")
    
    If (physFlag And &H8000&) Then
        Dim xLoc As DecalNet.IMessageMember
        Set xLoc = xPhysics.Struct("position")

        objEntity.Loc.Landblock = CLng(xLoc.Value("landcell"))
        objEntity.Loc.Xoff = CSng(xLoc.Value("x"))
        objEntity.Loc.Yoff = CSng(xLoc.Value("y"))
        objEntity.Loc.Zoff = CSng(xLoc.Value("z"))
    End If
    
    Dim sDebug As String
    
    sDebug = objEntity.Name
    
    If (gameFlags1 And &H8&) Then objEntity.Value = xGame.Value("value")
    If (gameFlags1 And &H40&) Then objEntity.MonarchID = xGame.Value("monarch")
    If (gameFlags1 And &H400&) Then objEntity.UsesLeft = xGame.Value("uses")
    If (gameFlags1 And &H800&) Then objEntity.TotalUses = xGame.Value("usesLimit")
    If (gameFlags1 And &H1000&) Then objEntity.StackCount = xGame.Value("stack")
    If (gameFlags1 And &H2000&) Then objEntity.StackMax = xGame.Value("stackLimit")
    If (gameFlags1 And &H4000&) Then objEntity.Container = xGame.Value("container")
    If (gameFlags1 And &H8000&) Then objEntity.Wielder = xGame.Value("equipper")
    If (gameFlags1 And &H10000) Then objEntity.Coverage = xGame.Value("equipPossible")
    If (gameFlags1 And &H20000) Then objEntity.Coverage2 = xGame.Value("equipActual")
    If (gameFlags1 And &H40000) Then objEntity.Coverage3 = xGame.Value("coverage")
    If (gameFlags1 And &H200000) Then objEntity.Burden = xGame.Value("burden")
    If (gameFlags1 And &H400000) Then objEntity.AssociatedSpellId = xGame.Value("spell")
    If (gameFlags1 And &H1000000) Then objEntity.Workmanship = xGame.Value("workmanship")
    
    'Material info
    If (gameFlags1 And &H80000000) Then
        objEntity.MaterialType = xGame.Value("material")

        'FIXME
        'Call g_ACConst.CheckNewMaterial(objEntity.MaterialType, objEntity.Name)
    End If
    
    'Now try to tell what kind of entity this is
    
    '----------------------------------------------------------------------------------------
    'Monsters/NPCs/Players...
    '----------------------------------------------------------------------------------------
    If (catFlag And &H10&) Then
    
        'myDebug "Mobile object of gameDataType: " & objEntity.gameDataType & " for catFlag: " & Hex(catFlag) & " behaveFlag: " & Hex(behaveFlag)
        
        'Players
        If (behaveFlag And &H8&) Then
            'myDebug "Object type PLAYER: " & objEntity.Name
            objEntity.ObjectType = TYPE_PLAYER
            
            'Set the PK Flag
            If (behaveFlag And &H2000000) Then
                objEntity.PlayerType = PLAYER_PINK  'PK-Lite
            ElseIf (behaveFlag And &H20&) Then
                objEntity.PlayerType = PLAYER_RED   'PK
            Else
                objEntity.PlayerType = PLAYER_WHITE 'NPK
            End If
            
            'Special case for the local player object which was created before the create object packet
            If objEntity.GUID = m_objPlayer.GUID Then
                Set m_objPlayer = objEntity
            Else
                If Not m_colPlayers.AddObject(objEntity) Then
                    myError "GameObjects.NetEcho_OnCreateObject: failed to add player " & objEntity.Name & " [" & objEntity.GUID & "] to Players collection"
                    'GoTo Fin
                End If
            End If

        ' Merchant
        ElseIf (behaveFlag And &H200&) Then
            'myDebug "Object type MERCHANT: " & objEntity.Name
            objEntity.ObjectType = TYPE_MERCHANT
            
            If Not m_colJunk.AddObject(objEntity) Then
                myError "GameObjects.NetEcho_OnCreateObject: failed to add merchant " & objEntity.Name & " [" & objEntity.GUID & "] to Junk collection"
                GoTo Fin
            End If
            
        'Monsters
        ElseIf (behaveFlag = &H14&) Then
            myDebug "NetEcho_OnCreateObject: Object type MONSTER: " & objEntity.Name
            objEntity.ObjectType = TYPE_MONSTER
            If Not m_colMonsters.AddObject(objEntity) Then
                myError "GameObjects.NetEcho_OnCreateObject: 1st try failed to add monster " & objEntity.Name & " [" & objEntity.GUID & "] to Monsters collection"
                If Not m_colMonsters.AddObject(objEntity) Then
                    myError "GameObjects.NetEcho_OnCreateObject: 2nd try failed to add monster " & objEntity.Name & " [" & objEntity.GUID & "] to Monsters collection"
                    'GoTo Fin
                End If
            End If
        
        'All the rest (NPCs, etc.)
        Else
            myDebug "GameObjects.NetEcho_OnCreateObject - Unhandled Mobile: " & objEntity.Name & " behaveFlag: " & Hex(behaveFlag)
            GoTo Fin
        End If
        
    '----------------------------------------------------------------------------------------
    'Items/Non living entities (can be picked up and can be selected )
    '----------------------------------------------------------------------------------------
    ElseIf (Not (behaveFlag And &H4&)) And (behaveFlag And &H10&) Then
        objEntity.ObjectType = TYPE_ITEM
        
        'Figure out what kind of item this is
        Dim lItemType As eItemTypes: lItemType = ITEM_UNKNOWN
        Dim bCanBePack As Boolean: bCanBePack = False
        
        If (catFlag And &H1&) Then
            lItemType = ITEM_MELEE_WEAPON
        ElseIf (catFlag And &H2&) Then
            lItemType = ITEM_ARMOR
            Call objEntity.UpdateArmorType
        ElseIf (catFlag And &H4&) Then
            lItemType = ITEM_CLOTHING
        ElseIf (catFlag And &H8&) Then
            lItemType = ITEM_JEWELRY
        ElseIf (catFlag And &H20&) Then
            lItemType = ITEM_FOOD
        ElseIf (catFlag And &H40&) Then
            lItemType = ITEM_PYREAL
        ElseIf (catFlag And &H80&) Then
            lItemType = ITEM_MISC
            If InStr(sObjName, "Foci") Then bCanBePack = True
        ElseIf (catFlag And &H100&) Then
            If InStr(LCase(sObjName), "arrow") Or InStr(LCase(sObjName), "quarrel") Then
                lItemType = ITEM_ARROW
            Else
                lItemType = ITEM_MISSILE_WEAPON
            End If
        ElseIf (catFlag And &H200&) Then
            lItemType = ITEM_CONTAINER
            bCanBePack = True
        ElseIf (catFlag And &H400&) Then
            lItemType = ITEM_BUNDLE
        ElseIf (catFlag And &H800&) Then
            lItemType = ITEM_GEM
        ElseIf (catFlag And &H1000&) Then
            lItemType = ITEM_COMPS
        ElseIf (catFlag And &H2000&) And InStr(LCase(sObjName), "scroll") Then
            lItemType = ITEM_SCROLL
        ElseIf (catFlag And &H2000&) Then
            lItemType = ITEM_BOOK
        ElseIf (catFlag And &H4000&) Then
            lItemType = ITEM_KEY
        ElseIf (catFlag And &H8000&) Then
            lItemType = ITEM_WAND
        ElseIf (catFlag And &H10000) Then
            lItemType = ITEM_PORTAL
        ElseIf (catFlag And &H40000) Then
            lItemType = ITEM_TRADENOTE
        ElseIf (catFlag And &H80000) Then
            lItemType = ITEM_MANA_STONES
        ElseIf (catFlag And &H200000) Then
            lItemType = ITEM_PLANT
        ElseIf (catFlag And &H400000) Then
            lItemType = ITEM_BASE_COOKING
        ElseIf (catFlag And &H800000) Then
            'lItemType = ITEM_BASE_ALCHEMY
            lItemType = ITEM_BASE_FLETCHING
        ElseIf (catFlag And &H1000000) Then
            lItemType = ITEM_BASE_FLETCHING
        ElseIf (catFlag And &H2000000) Then
            lItemType = ITEM_CRAFTED_COOKING
        ElseIf (catFlag And &H4000000) Then
            lItemType = ITEM_CRAFTED_ALCHEMY
        ElseIf (catFlag And &H8000000) Then
            lItemType = ITEM_CRAFTED_FLETCHING
        ElseIf (catFlag And &H20000000) Then
            lItemType = ITEM_UST
        ElseIf (catFlag And &H40000000) Then
            lItemType = ITEM_SALVAGE
        End If
        
        'ObjFlag2 overrides ObjFlag1
        If (behaveFlag And &H1000&) Then
            lItemType = ITEM_DOOR
        ElseIf (behaveFlag And &H2000&) Then
            lItemType = ITEM_CORPSE
        ElseIf (behaveFlag And &H4000&) Then
            lItemType = ITEM_LIFESTONE
        ElseIf (behaveFlag And &H8000&) Then
            lItemType = ITEM_FOOD
        ElseIf (behaveFlag And &H10000) Then
            lItemType = ITEM_HEALING_KIT
        ElseIf (behaveFlag And &H20000) Then
            lItemType = ITEM_LOCKPICK
        ElseIf (behaveFlag And &H40000) Then
            lItemType = ITEM_PORTAL
        ElseIf (behaveFlag And &H800000) Then
            lItemType = ITEM_FOCI
            bCanBePack = True
        ElseIf (behaveFlag And &H1&) Then 'must be last
            lItemType = ITEM_CONTAINER
            bCanBePack = True
        End If
        
        objEntity.ItemType = lItemType
        objEntity.IsPack = bCanBePack And (objEntity.Container = m_objPlayer.GUID)
        
        myDebug "Name: " & objEntity.Name & " ItemType: " & objEntity.ItemType & " catFlag: " & Hex(catFlag) & " behaveFlag: " & Hex(behaveFlag)
        
        If (objEntity.Wielder = m_objPlayer.GUID) And (Not objEntity.IsPack) Then
            objEntity.Equiped = True
        End If
        
        'If it's a cast spell, just bail here, as we don't care about it
        If (objEntity.ItemType = ITEM_UNKNOWN) And (gameFlags1 And &H400000) Then
            'myDebug "Name: " & objEntity.Name & " appears to be a spell, NOT adding to inventory"
            GoTo Fin
        End If
        
        'Check to see if we put this item in World or Inventory
        'If located in main inventory pack
        'or item is equiped by us,
        'or item belongs to one of our inventory packs, add it to inventory - else it belongs to world
        If (objEntity.Container = m_objPlayer.GUID) _
        Or (objEntity.Wielder = m_objPlayer.GUID) _
        Or m_Items.InInventory(objEntity.Container) Then
            If m_Items.AddInventoryItemEx(objEntity) Then
                Call m_Items.UpdateInventoryCount
            Else
                myError "GameObjects.NetEcho_OnCreateObject: failed to add item " & objEntity.Name & " [" & objEntity.GUID & "] to Inventory Items collection"
                'GoTo Fin
            End If
        Else
            If Not m_Items.AddWorldItemEx(objEntity) Then
                myError "GameObjects.NetEcho_OnCreateObject: failed to add item " & objEntity.Name & " [" & objEntity.GUID & "] to World Items collection"
                'GoTo Fin
            End If
        End If
        
    Else
        GoTo Fin
    End If
    
    'Fire the object creation event
    RaiseEvent OnCreateObject(objEntity)

Fin:
    Set objEntity = Nothing
    'myDebug "OnCreateObject done"
    Exit Sub
ErrorHandler:
    myError "GameObjects.NetEvent_OnCreateObject - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Friend Sub NetEvent_OnIdentifyObject(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim lObjGUID As Long
    Dim lAssessFlag As Long
    Dim lCharFlags As Long
    Dim iHealth As Integer, iMaxHealth As Integer
    Dim iStam As Integer, iMaxStam As Integer
    Dim iMana As Integer, iMaxMana As Integer
    Dim sMonarch As String
    Dim objEntity As acObject
    Dim iNumData As Integer, i As Integer
    Dim sVal As String, lVal As Long, lKey As Long, dVal As Double, bVal As Boolean
    Dim bAddImbueComma As Boolean
    Dim bCrushBAdded  As Boolean
    Dim bBSAdded As Boolean
    Dim bRCAdded As Boolean
    Dim isNewObj As Boolean
    
    bAddImbueComma = False
    bCrushBAdded = False
    bBSAdded = False
    bRCAdded = False
        
    lObjGUID = pMsg.Value("object")
    lAssessFlag = pMsg.Value("flags")
    lCharFlags = pMsg.Value("flags1")

    Set objEntity = FindObject(lObjGUID, False)
    
    If Not Valid(objEntity) Then
        myError "OnIdentifyObject: FindObject returned Nothing for lObjGUID: " & lObjGUID
        myDebug "OnIdentifyObject: not found! Creating new acObject for lObjGUID: " & lObjGUID
        Set objEntity = New acObject
        objEntity.GUID = lObjGUID
        isNewObj = True
        'GoTo Fin
    Else
        isNewObj = False
        If GetModObject(lObjGUID, objEntity) Then
            objEntity.LastIDTime = g_Time
            myDebug "OnIdentifyObject: WILL UPDATE (" & g_Time & ") " & objEntity.Name & " : " & lObjGUID & " -- AssessFlag=" & Hex(lAssessFlag) & " -- lCharFlags=" & Hex(lCharFlags)
        End If
    End If

    'myDebug "OnIdentifyObject: " & objEntity.Name & " : " & lObjGUID & " -- AssessFlag=" & Hex(lAssessFlag) & " -- lCharFlags=" & Hex(lCharFlags)
    
    'Remember when we last IDed this item
    objEntity.LastIDTime = g_Time

    '----------------------------------------------------------------
    ' Strings (description, useInstructions, etc)
    '----------------------------------------------------------------
1    If (lAssessFlag And &H8&) Then

        Dim xS As DecalNet.IMessageMember
        iNumData = pMsg.Value("stringCount")
        Set xS = pMsg.Struct("strings")

        'myDebug "IDObject: Strings: " & objEntity.Name & " [NumData=" & iNumData & "]"

        For i = 0 To iNumData - 1
            sVal = xS.Struct(i).Value("value")
            lKey = xS.Struct(i).Value("key")
        
            myDebug "IDObject: Strings: " & objEntity.Name & " Key:" & Hex(lKey) & " Value:" & sVal
            
            Select Case lKey
                Case &H1&  'Name
                    If (objEntity.Name = "UnknownEntity") Or (objEntity.Name = "LocalPlayer") Or (objEntity.Name = "None") Or isNewObj Then
                        objEntity.Name = sVal
                        myDebug "New Obj Name: " & sVal
                    End If
                
                Case &H7& 'Inscription
                    objEntity.Inscription = sVal

                Case &H8& 'Inscriber
                    objEntity.Inscriber = sVal

                Case &H10& 'Description -- That's where we can find "Killed by PlayerName." kind of info
                    'MyDebug "ID: " & objEntity.Name & " - Description : " & sVal
                    objEntity.Description = sVal

                    If objEntity.ItemType = ITEM_CORPSE Then
                        If InStr(LCase(sVal), LCase(m_objPlayer.Name)) Then
                            objEntity.KillerName = m_objPlayer.Name
                        ElseIf InStr(LCase(sVal), "killed by ") Then
                            sVal = Mid(sVal, 1, Len(sVal) - 1)  'remove the dot at the end
                            objEntity.KillerName = Trim(Mid(sVal, Len("killed by ") + 1))
                            'MyDebug "ID: " & objEntity.Name & " - Killer : " & objEntity.KillerName
                        End If
                    End If
                
                Case &HE&  'Use Instruction
                    objEntity.UsageInstructions = sVal
                    'myDebug "ID: " & objEntity.Name & " - Use Instruction : " & sVal

                Case &HF&  'Simple Description
                    objEntity.ShortDesc = sVal
                    'myDebug "ID: " & objEntity.Name & " - Simple Description : " & sVal

                Case &H2B& 'Date of Birth
                    If isNewObj Then
                        objEntity.ObjectType = TYPE_PLAYER
                    End If
                    
            End Select
        Next i

    End If
    
    '----------------------------------------------------------------
    ' General item stats information
    '----------------------------------------------------------------
2    If (lAssessFlag And &H1&) Then
        Dim xDwords As DecalNet.IMessageMember
        Dim xVal As DecalNet.IMessageMember
        
        iNumData = pMsg.Value("dwordCount")
        Set xDwords = pMsg.Struct("dwords")

        'myDebug "IDObject: Asses Item General: " & objEntity.Name & " [NumData=" & iNumData & "]"

        For i = 0 To iNumData - 1
            Set xVal = xDwords.Struct(i)
            lVal = xVal.Value("value")
            lKey = xVal.Value("key")
            
            myDebug "IDObject: H1: " & objEntity.Name & " Key:" & Hex(lKey) & " Val:" & lVal
            
            Select Case lKey
                Case &H5& 'Burden
                    objEntity.Burden = lVal

                Case &H11& 'Rare #
                    objEntity.IsRare = True
                    objEntity.RareNumber = lVal
                    myDebug "IDObject: Rare #: " & objEntity.Name & " Key:" & Hex(lKey) & " Rare #:" & lVal
                    'myError "IDObject: Rare #: " & objEntity.Name & " Key:" & Hex(lKey) & " Rare #:" & lVal
                    RaiseEvent OnRareDetected(objEntity)
                
                Case &H13& 'Value
                    objEntity.Value = lVal

                Case &H1C& 'Armor level
                    objEntity.ArmorLevel = lVal
                    myDebug "IDObject: " & objEntity.Name & " - Armor Level : " & objEntity.ArmorLevel
                    If objEntity.ItemType <> ITEM_ARMOR Then
                        myDebug "IDObject: " & objEntity.Name & " - NOT Armor: " & objEntity.ItemType
                        'objEntity.ItemType = ITEM_ARMOR
                        Call objEntity.UpdateArmorType
                    End If

                Case &H21&  'Bonded
                    'objEntity.Bonded = True
                    
                Case &H24&  'Unechantable
                    objEntity.unEnchantable = True

                Case &H2D&  'Elemental Damage Bonus: DamageType
                    'objEntity.ElementBonusDamageType = lVal
                    
                Case &H56&  'Min level restriction
                    'objEntity.minLevel = lVal

                Case &H57&  'Max level restriction
                    'objEntity.maxLevel = lVal

                Case &H5B&  'Total Uses
                    'objEntity.totalUses = lVal

                Case &H5C&  'Uses Remaining
                    'objEntity.usesRemaining = lVal

                Case &H69& 'Workmanship
                    objEntity.Workmanship = lVal

                Case &H5B& 'Maximum Uses
                    objEntity.TotalUses = lVal

                Case &H5C& 'Current Uses
                    objEntity.UsesLeft = lVal
                    
                Case &H6A& 'Spellcraft
                    objEntity.Spellcraft = lVal
                    
                Case &H6B& 'Mana Remaining
                    objEntity.Mana = lVal

                Case &H6C& 'Max Mana
                    objEntity.MaxMana = lVal

                Case &H6D& 'Arcane Lore req
                    objEntity.LoreReq = lVal
                
                Case &H6E& 'Rank activation req
                    objEntity.RankReq = lVal

                Case &H83& 'Material Type
                    'Call g_ACConst.CheckNewMaterial(objEntity.MaterialType, objEntity.Name)
                    objEntity.MaterialType = lVal

                Case &H9E& 'Wield Req Type
                    'MyDebug "ID: " & objEntity.Name & " - Wield Req Type : " & objEntity.WieldReqType
                    objEntity.WieldReqType = lVal

                Case &H9F& ' WieldRequirement - ID
                    'MyDebug "ID: " & objEntity.Name & " - Wield Req Id : " & objEntity.WieldReqId & " = " & g_ACConst.GetWeaponSkillName(objEntity.WieldReqId)
                    objEntity.WieldReqId = lVal

                Case &HA0& 'WieldRequirement - Amount
                    'MyDebug "ID: " & objEntity.Name & " - Wield Req Value : " & objEntity.WieldReqVal
                    objEntity.WieldReqVal = lVal

                Case &HA6& 'Slayer Species
                    objEntity.slayerType = lVal

                Case &HAB& 'Number times tinked
                    objEntity.TinkCount = lVal

                Case &HAA& 'Workmanship divisor
                    If lVal <> 0 Then
                        'myDebug "ID: " & objEntity.Name & " - Workmanship Divisor: " & lVal & " - Updated Workmanship = " & objEntity.Workmanship
                        objEntity.Workmanship = Round(objEntity.Workmanship / lVal)
                    End If

                Case &HB0& 'Activation Skill Name
                    objEntity.ActivateSkill = g_Const.GetSkillName(lVal)
                
                Case &H73& 'Activation Skill level
                    objEntity.ActivateSkillVal = lVal
                
                Case &HBC& 'Racial Activation Req
                    objEntity.RaceReq = ToRaceString(lVal)
                
                Case &HB3& 'special properties
                    '0x01=Critical Strike, 0x02=Crippling Blow, 0x04=Armor Rending, 0x80000000= Phantasmal
                    myDebug "IDObject: H1->B3: " & objEntity.Name & " Key:" & Hex(lKey) & " Val:" & lVal
                    myDebug "IDObject: Hex(Imbue): " & Hex(lVal)
                    
                    'These need to be a bitwise AND as they can exist with other imbues
                    '0x20000000= Magic Absorbing
                    '0x20000004= Magic Absorbing (really some other imbue + Magic Absorbing
                    If (lVal And &H20000000) Then
                        lVal = Hex(lVal) - &H20000000
                        If InStr(objEntity.Imbue, "Magic Absorbing") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Magic Absorbing")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Magic Absorbing")
                                bAddImbueComma = True
                            End If
                        End If
                    End If
                    '0x80000000= Phantasmal
                    If (lVal And &H80000000) Then
                        lVal = Hex(lVal) - &H80000000
                        If InStr(objEntity.Imbue, "Phantasmal") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Phantasmal")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Phantasmal")
                                bAddImbueComma = True
                            End If
                        End If
                    End If
                    
                    If lVal = 1 Then
                        If InStr(objEntity.Imbue, "Critical Strike") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Critical Strike")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Critical Strike")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 2 Then
                        If InStr(objEntity.Imbue, "Crippling Blow") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Crippling Blow")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Crippling Blow")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 4 Then
                        If InStr(objEntity.Imbue, "Armor Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Armor Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Armor Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 8 Then
                        objEntity.DamageType = DMG_SLASHING
                        If InStr(objEntity.Imbue, "Slash Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Slash Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Slash Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 16 Then
                        objEntity.DamageType = DMG_PIERCING
                        If InStr(objEntity.Imbue, "Pierce Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Pierce Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Pierce Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 32 Then
                        objEntity.DamageType = DMG_BLUDGEONING
                        If InStr(objEntity.Imbue, "Bludgeon Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Bludgeon Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Bludgeon Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 64 Then
                        objEntity.DamageType = DMG_ACID
                        If InStr(objEntity.Imbue, "Acid Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Acid Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Acid Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 128 Then
                        objEntity.DamageType = DMG_COLD
                        If InStr(objEntity.Imbue, "Frost Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Frost Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Frost Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 256 Then
                        objEntity.DamageType = DMG_LIGHTNING
                        If InStr(objEntity.Imbue, "Lightning Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Lightning Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Lightning Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    ElseIf lVal = 512 Then
                        objEntity.DamageType = DMG_FIRE
                        If InStr(objEntity.Imbue, "Fire Rending") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Fire Rending")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Fire Rending")
                                bAddImbueComma = True
                            End If
                        End If
                    End If
                    
                
                Case &HCC& 'ElementBonusDamage
                    objEntity.ElementBonusDamage = lVal
                    'myDebug "ID: " & objEntity.Name & " - Element Bonus Damage: " & objEntity.ElementBonusDamage
                    
                Case &H107& 'Resistance Cleaving Type
                    If lVal = 1 Then
                        objEntity.ResistanceCleavingType = "Slashing"
                    End If
                    
                    If lVal = 2 Then
                        objEntity.ResistanceCleavingType = "Piercing"
                    End If
                
                'Case Else
                '    myDebug "IDObject: H1: Else unknown prop: " & objEntity.Name & " Key:" & Hex(lKey) & " Val:" & lVal

            End Select
        Next i
    End If

    '----------------------------------------------------------------
    ' Booleans
    '----------------------------------------------------------------
3    'If (lAssessFlag And &H2&) Then

        'Dim xB As DecalNet.IMessageMember
        'iNumData = pMsg.Value("booleanCount")
        'Set xB = pMsg.Struct("booleans")

        'myDebug "IDObject: Booleans: " & objEntity.Name & " [NumData=" & iNumData & "]"

        'For i = 0 To iNumData - 1
        '    bVal = xB.Struct(i).Value("value")
        '    lKey = xB.Struct(i).Value("key")
        '
        '    'myDebug "Boolean"
        '    myDebug "IDObject: Booleans:  " & objEntity.Name & " Key:" & Hex(lKey) & " Val:" & bVal
        '
        '    Select Case lKey
        '        'Case &H5B& 'Retained
        '            'objEntity.retained = True
        '        'Case &H63& 'Ivoryable
        '            'objEntity.ivoryable = True
        '        'Case &H64& 'Dyable
        '            'objEntity.dyable = True
        '        'Case &HE&  'Use Instruction
        '            'myDebug "ID: " & objEntity.Name & " - Use Instruction : " & sVal
        '        'Case &HF3&  'Unlimited Uses
        '            'myDebug "ID: " & objEntity.Name & " - Unlimited Uses : " & sVal
        '    End Select
        'Next i

    'End If
    
    '----------------------------------------------------------------
    ' Wand/Item Bonus Info
    '----------------------------------------------------------------
4    If (lAssessFlag And &H4&) Then

        Dim xD As DecalNet.IMessageMember
        iNumData = pMsg.Value("doubleCount")
        Set xD = pMsg.Struct("doubles")

        For i = 0 To iNumData - 1
            dVal = xD.Struct(i).Value("value")
            lKey = xD.Struct(i).Value("key")
            
            myDebug "IDObject: H4: " & objEntity.Name & " Key:" & Hex(lKey) & " Val:" & dVal
            
            Select Case lKey

                Case &H1D  'Melee Defense Bonus
                    objEntity.DefenseBonus = Round((dVal - 1) * 100)
                    If objEntity.DefenseBonus < 0 Then objEntity.DefenseBonus = 0
                    'myDebug "ID: " & objEntity.Name & " - MeleeDef Mod (dVal:" & dVal & ") = +" & objEntity.DefenseBonus & "%"
                
                Case &H90 'Mana Conversion modifier
                    objEntity.ManaConvMod = Round(dVal * 100)
                    If objEntity.ManaConvMod < 0 Then objEntity.ManaConvMod = 0
                    'MyDebug "ID: " & objEntity.Name & " - Mana Conv Mod (dVal:" & dVal & ") = +" & objEntity.ManaConvMod & "%"
                
                Case &H95   'Missile Def Bonus
                    objEntity.MissileDefense = Round((dVal - 1) * 100)
                    If objEntity.MissileDefense < 0 Then objEntity.MissileDefense = 0
                    'myDebug "ID: " & objEntity.Name & " - MissileDef Mod (dVal:" & dVal & ") = +" & objEntity.MissileDefense & "%"
                
                Case &H96   'Magic Def Bonnus
                    objEntity.MagicDefense = Round((dVal - 1) * 100, 1)
                    If objEntity.MagicDefense < 0 Then objEntity.MagicDefense = 0
                    'myDebug "ID: " & objEntity.Name & " - MagicDefense Mod (dVal:" & dVal & ") = +" & objEntity.MagicDefense & "%"
                
                Case &H98   'PvM bonus
                    objEntity.PvMBonus = Round(dVal * 100) - 100
                    If objEntity.PvMBonus < 0 Then objEntity.PvMBonus = 0
                    myDebug "ID: " & objEntity.Name & " - PvM Bonus (dVal:" & dVal & ") = +" & objEntity.PvMBonus
                    
                Case &H88 'Crushing Blow
                    If lVal >= 1 Then
                        myDebug "Found Crushing Blow property."
                        objEntity.CrushingBlow = True
                        If InStr(objEntity.Imbue, "Crushing Blow") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Crushing Blow")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Crushing Blow")
                                bAddImbueComma = True
                            End If
                        End If
                    End If
                
                Case &H93 'Biting Strike
                    If lVal >= 1 Then
                        myDebug "Found Biting Strike property."
                        objEntity.BitingStrike = True
                        If InStr(objEntity.Imbue, "Biting Strike") = 0 Then
                            If bAddImbueComma Then
                                objEntity.Imbue = LTrim(objEntity.Imbue & ", Biting Strike")
                            Else
                                objEntity.Imbue = LTrim(objEntity.Imbue & "Biting Strike")
                                bAddImbueComma = True
                            End If
                        End If
                    End If
                    
                Case &H9D 'Resistance Cleaving
                    If lVal >= 1 Then
                        If InStr(objEntity.ResistanceCleaving, "Resistance Cleaving:") = 0 And (bRCAdded = False) Then
                            If bAddImbueComma Then
                                objEntity.ResistanceCleaving = LTrim(objEntity.ResistanceCleaving & ", Resistance Cleaving: ")
                                bRCAdded = True
                            Else
                                objEntity.ResistanceCleaving = LTrim(objEntity.ResistanceCleaving & "Resistance Cleaving: ")
                                bAddImbueComma = True
                                bRCAdded = True
                            End If
                        End If
                    End If
            End Select
        Next i
    End If


    '----------------------------------------------------------------
    ' Spells on Item
    '----------------------------------------------------------------
5    If (lAssessFlag And &H10&) Then
        
        Dim lSpellID As Long
        Dim sSpellName As String
        Dim xSpells As DecalNet.IMessageMember

        iNumData = pMsg.Value("spellCount")
        Set xSpells = pMsg.Struct("spells")
        
        'myDebug "IDObject: Spells: " & objEntity.Name & " [NumData=" & iNumData & "]"
        
        'If iNumData > 0 Then
        '    Set objEntity.SpellsActive = New Dictionary
        'End If
        
        For i = 0 To iNumData - 1
            lSpellID = xSpells.Struct(i).Value("spell")
            sSpellName = g_Const.GetSpellName(lSpellID)
            
            If sSpellName <> "" Then
                Dim sFlags As Integer
                sFlags = xSpells.Struct(i).Value("flags")
                
                'If Not objEntity.Spells.Exists(sSpellName) Then
                '    Call objEntity.Spells.Add(sSpellName, sFlags)
                'End If
                
                If sFlags = 0 Then
                    If Not objEntity.Spells.Exists(sSpellName) Then
                        Call objEntity.Spells.Add(sSpellName, sFlags)
                    End If
                Else
                    If Not objEntity.SpellsActive.Exists(sSpellName) Then
                        Call objEntity.SpellsActive.Add(sSpellName, sFlags)
                    End If
                End If
                
                myDebug "ID: " & objEntity.Name & " - Spell: " & sSpellName & "  flags: " & sFlags
                
                If InStr(LCase(sSpellName), "epic") Then
                    'PrintMessage "[Major] Item Detected : " & objEntity.Name & " - " & sSpellName
                    objEntity.HasEpics = True
                    RaiseEvent OnEpicDetected(objEntity, sSpellName)
                ElseIf InStr(LCase(sSpellName), "major") Then
                    'PrintMessage "[Major] Item Detected : " & objEntity.Name & " - " & sSpellName
                    objEntity.HasMajors = True
                    RaiseEvent OnMajorDetected(objEntity, sSpellName)
                ElseIf InStr(LCase(sSpellName), "minor") Then
                    'PrintMessage "[Minor] Item Detected : " & objEntity.Name & " - " & sSpellName
                    objEntity.HasMinors = True
                    RaiseEvent OnMinorDetected(objEntity, sSpellName)
                'ElseIf InStr(LCase(sSpellName), "prodigal") Then
                '    myDebug "IDObject: Spells: " & objEntity.Name & " Spellname:" & sSpellName
                '    objEntity.IsRare = True
                '    RaiseEvent OnRareDetected(objEntity)
                End If
            Else
                myDebug "ID: " & objEntity.Name & " - Spell : Unknown spell " & lSpellID
            End If

        Next i
    End If

    '----------------------------------------------------------------
    ' Weapon Info
    '----------------------------------------------------------------
6    If (lAssessFlag And &H20&) Then
        'myDebug "OnIdentifyObject H20 Weap: " & objEntity.Name & " : lAssessFlag: " & Hex(lAssessFlag) & "  lCharFlags: " & Hex(lCharFlags)
        
        objEntity.DamageFlags = pMsg.Value("weapType")
        objEntity.SkillUsed = pMsg.Value("weapSkill")
        objEntity.HighDamage = pMsg.Value("weapDamage")
        objEntity.Variance = pMsg.Value("weapVariance")
        
        objEntity.AttackBonus = Round((pMsg.Value("weapAttack") - 1) * 100)
        objEntity.DamageModifier = Round((pMsg.Value("weapModifier") - 1) * 100)

        If objEntity.AttackBonus < 0 Then objEntity.AttackBonus = 0
        If objEntity.DamageModifier < 0 Then objEntity.DamageModifier = 0

        Select Case objEntity.DamageFlags
            Case &H1&  'Slashing
                objEntity.DamageType = DMG_SLASHING
            Case &H2&  'Pierce
                objEntity.DamageType = DMG_PIERCING
            Case &H4&  'Bludge
                objEntity.DamageType = DMG_BLUDGEONING
            Case &H8&  'Cold
                objEntity.DamageType = DMG_COLD
            Case &H10& 'Fire
                objEntity.DamageType = DMG_FIRE
            Case &H20& 'Acid
                objEntity.DamageType = DMG_ACID
            Case &H40& 'Lightning
                objEntity.DamageType = DMG_LIGHTNING
        End Select

        'MyDebug "ID : Weapon Info - " & objEntity.Name & " : Type=" & objEntity.GetWeaponTypeName & ", HighDmg=" & objEntity.HighDamage & ", LowDmg=" & objEntity.GetLowDamage & " Var=" & objEntity.Variance & ", Att%=" & objEntity.AttackBonus & ", Def%=" & objEntity.DefenseBonus & ", DmgFlags=" & objEntity.DamageFlags
        'If objEntity.DamageModifier <> 0 Then MyDebug "Damage Mod : +" & objEntity.DamageModifier & "%"
        
    End If
    
    'If (lAssessFlag And &H103D&) Then ' Cast on Strike property
    If Hex(lAssessFlag) = &H103D& Then ' Cast on Strike property
        myDebug "OnIdentifyObject H20 -> H103D Weap: " & objEntity.Name & " : lAssessFlag: " & Hex(lAssessFlag) & "  lCharFlags: " & Hex(lCharFlags)
        If InStr(objEntity.Imbue, "Cast on Strike") = 0 Then
            If bAddImbueComma Then
                objEntity.Imbue = LTrim(objEntity.Imbue & ", Cast on Strike")
            Else
                objEntity.Imbue = LTrim(objEntity.Imbue & "Cast on Strike")
                bAddImbueComma = True
            End If
        End If
    End If
    

    '----------------------------------------------------------------
    ' Character Assess (Player)
    '----------------------------------------------------------------
7    If (lAssessFlag And &H100&) Then

        myDebug "OnIdentifyObject H100: " & objEntity.Name & " : lAssessFlag: " & Hex(lAssessFlag) & "  lCharFlags: " & Hex(lCharFlags)

        If isNewObj And (objEntity.ObjectType <> TYPE_PLAYER) Then
            objEntity.ObjectType = TYPE_MONSTER
        End If

        'If (lCharFlags And &H4) Then
            'MyDebug "OnIdentifyObject: &H4&"
            iHealth = pMsg.Value("health")
            iMaxHealth = pMsg.Value("healthMax")

            If iHealth > 0 Then objEntity.Health = iHealth
            If iMaxHealth > 0 Then objEntity.MaxHealth = iMaxHealth
            
            'myDebug objEntity.Name & "'s Health : " & objEntity.Health & "/" & objEntity.MaxHealth

            'objEntity.Level = pMsg.Value("level")
        'End If

8        If (lCharFlags And &H8&) Then
            'Stamina
            iStam = pMsg.Value("stamina")
            iMaxStam = pMsg.Value("staminaMax")

            If iStam > 0 Then objEntity.Stamina = iStam
            If iMaxStam > 0 Then objEntity.MaxStamina = iMaxStam
            'MyDebug objEntity.Name & "'s Stamina : " & objEntity.Stamina & "/" & objEntity.MaxStamina

            'Mana
            iMana = pMsg.Value("mana")
            iMaxMana = pMsg.Value("manaMax")

            If iMana > 0 Then objEntity.Mana = iMana
            If iMaxMana > 0 Then objEntity.MaxMana = iMaxMana
            'MyDebug objEntity.Name & "'s Mana : " & objEntity.Mana & "/" & objEntity.MaxMana

            With objEntity
                .AttribStrenght = pMsg.Value("strength")
                .AttribEndurance = pMsg.Value("endurance")
                .AttribCoordination = pMsg.Value("coordination")
                .AttribQuickness = pMsg.Value("quickness")
                .AttribFocus = pMsg.Value("focus")
                .AttribSelf = pMsg.Value("self")
            End With
        End If

        'If (lCharFlags And &H2&) Then
        '    With objEntity
        '        .MonarchName = pMsg.Value("monarch")
        '        .Followers = pMsg.Value("followers")
        '        .Leadership = pMsg.Value("leadership")
        '        .Gender = pMsg.Value("gender")
        '        .Race = pMsg.Value("race")
        '        .Class = pMsg.Value("class")
        '        .FellowshipName = pMsg.Value("fellowship")
        '    End With
        'End If

    End If 'lAssesFlag & H100
        
    'If (objEntity.Container = m_objPlayer.GUID) _
    'Or (objEntity.Wielder = m_objPlayer.GUID) _
    'Or m_Items.InInventory(objEntity.Container) Then
    '    m_Items.AddInventoryItemEx (objEntity)
    'Else
    '    m_Items.AddWorldItemEx (objEntity)
    'End If
    
    ' Do we need to add it?
    'If isNewObj Then
    '    If m_Items.AddWorldItemEx(objEntity) Then
    '        myError "GameObjects.NetEcho_OnIdentifyObject: added new item to World: " & objEntity.Name
    '        myDebug "GameObjects.NetEcho_OnIdentifyObject: added new item to World: " & objEntity.Name
    '    Else
    '        myError "GameObjects.NetEcho_OnIdentifyObject: failed to add NEW item " & objEntity.Name & " [" & objEntity.GUID & "] to World Items collection"
    '    End If
    '    RaiseEvent OnCreateObject(objEntity)
    'End If

    'Trigger the assess event
    RaiseEvent OnIdentifyObject(objEntity)

Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "GameObjects.NetEvent_OnIdentifyObject - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

Friend Sub NetEvent_OnSetPosition(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim lObjGUID As Long
    Dim objEntity As acObject
    Dim lFlag As Long

    lObjGUID = pMsg.Value("object")

    Set objEntity = FindObject(lObjGUID, False)
    If Valid(objEntity) Then
        Dim lLandblock As Long
        Dim oldLoc As acLoc
        Dim x As DecalNet.IMessageMember
        
        Set x = pMsg.Struct("position")
        lLandblock = x.Value("landcell")
            
        Set oldLoc = objEntity.Loc.Clone

        'AC sometimes send a bugged empty landblock - ignore it
        If lLandblock <> 0 Then
            objEntity.Loc.Landblock = lLandblock
        End If
        objEntity.Loc.Xoff = x.Value("x")
        objEntity.Loc.Yoff = x.Value("y")
        objEntity.Loc.Zoff = x.Value("z")
        
        If Not objEntity.Loc.Equals(oldLoc) Then
            RaiseEvent OnObjectMoved(objEntity)
            objEntity.timeData = g_Time
        End If
            
            'f = p.Value("flags")
            'If ((f And &H8) = 0) Then
            '    wQuat = p.Value("wQuat")
            'Else
            '    wQuat = 0#
            'End If
            'If ((f And &H10) = 0) Then
            '    xQuat = p.Value("xQuat")
            'Else
            '    xQuat = 0#
            'End If
            'If ((f And &H20) = 0) Then
            '    yQuat = p.Value("yQuat")
            'Else
            '    yQuat = 0#
            'End If
            'If ((f And &H40) = 0) Then
            '    zQuat = p.Value("zQuat")
            'Else
            '    zQuat = 0#
            'End If
            'If ((f And &H1) <> 0) Then
            '    dx = p.Value("dx")
            '    dy = p.Value("dy")
            '    dz = p.Value("dz")
            'Else
            '    dx = 0#
            '    dy = 0#
            '    dz = 0#
            'End If
            
            
    End If

Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "OnSetPosition"
    Resume Fin
End Sub

'NetEvent_OnSetCharPosition
Friend Sub NetEvent_OnSetCharPosition(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    '
    'Worthless Network packet, as it's only used for last corpse location
    '
    
    Dim lFlag As Long

    Dim lLandblock As Long
    Dim x As DecalNet.IMessageMember
        
    Set x = pMsg.Struct("position0")
    lLandblock = x.Value("landcell")
        
            'objEntity.Loc.Landblock = lLandblock
            'objEntity.Loc.Xoff = x.Value("x")
            'objEntity.Loc.Yoff = x.Value("y")
            'objEntity.Loc.Zoff = x.Value("z")

Fin:
    Exit Sub
ErrorHandler:
    myError "OnSetCharPosition"
    Resume Fin
End Sub

'NetEvent_OnApproachVendor
Friend Sub NetEvent_OnApproachVendor(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim lObjGUID As Long
    Dim objEntity As acObject
    Dim maxBuy As Long
    Dim fractBuy As Single
    Dim fractSell As Single

    lObjGUID = pMsg.Value("merchant")

    Set objEntity = FindObject(lObjGUID, False)
    
    If Valid(objEntity) Then
    
        Set m_objVendor = objEntity
        Set m_colVendorInv = New colObjects
        m_colVendorInv.Description = "Vendor Inventory"
        
        maxBuy = pMsg.Value("buyValue")
        fractBuy = pMsg.Value("buyRate")
        fractSell = pMsg.Value("sellRate")
        
        objEntity.VendorMaxBuy = maxBuy
        objEntity.VendorFractBuy = fractBuy
        objEntity.VendorFractSell = fractSell
        
        myDebug "Vendor: maxBuy:" & objEntity.VendorMaxBuy & " fractBuy: " & objEntity.VendorFractBuy & " fractSell: " & objEntity.VendorFractSell
        
        'update the vendor inventory
        Dim iObjGUID As Long
        Dim invEntity As acObject
        Dim lAssessFlag As Long
        Dim lCharFlags As Long
        Dim iNumData As Integer, i As Integer
        Dim sVal As String, lVal As Long, dVal As Double
        Dim x As DecalNet.IMessageMember
        
        iNumData = pMsg.Value("itemCount")

        For i = 0 To iNumData - 1
            Set invEntity = New acObject
            
            invEntity.GUID = pMsg.Struct("items").Struct(i).Value("object")
            
            Set x = pMsg.Struct("items").Struct(i).Struct("game")
            
            invEntity.Name = x.Value("name")
            invEntity.Icon = x.Value("icon")
            
            lVal = x.Value("flags1")
            ' Value
            If (lVal And &H8&) Then invEntity.Value = x.Value("value")
            
            If (lVal And &H1000) Then objEntity.StackCount = x.Value("stack")
            If (lVal And &H2000) Then objEntity.StackMax = x.Value("stackLimit")
            If (lVal And &H200000) Then objEntity.Burden = x.Value("burden")
            'If (lVal And &H8000) Then objEntity.Wielder = x.Value("equipper")
            'If (lVal And &H1000000) Then objEntity.Workmanship = x.Value("workmanship")
                
            'Call m_colVendorInv.Add(invEntity.GUID, invEntity)
            Call m_colVendorInv.AddObject(invEntity)
            
            'myDebug "Vendor Inventory: guid: " & invEntity.GUID & " name: " & invEntity.Name & " flags2: " & lVal & " Val: " & invEntity.Value
        Next i
        
        
        'Trigger event
        RaiseEvent OnApproachVendor(objEntity, m_colVendorInv)
    Else
        myDebug "OnApproachVendor: could not find valid Vendor object: " & lObjGUID
    End If

Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "OnApproachVendor -- Err: " & Err.Description
    Resume Fin
End Sub

''Force vuln/enabled info reupdate
'Public Sub UpdateMonstersInfo()
'On Error GoTo ErrorHandler
'    Dim objMonster As acObject
'
'    For Each objMonster In m_colMonsters
'        Call objMonster.GetMonsterInfo
'    Next objMonster
'
'fin:
'    Set objMonster = Nothing
'    Exit Sub
'ErrorHandler:
'    myError "GameEntities.UpdateMonstersInfo"
'    Resume fin
'End Sub


'#####################################################################################
'#
'#                              Other Methods
'#
'#####################################################################################

Friend Sub SetCurrentSelection(ByVal lObjGUID As Long)
On Error GoTo ErrorHandler

    Set m_objSelected = FindObject(lObjGUID)
    RaiseEvent OnSelectObject(m_objSelected)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "GameObjects.SetCurrentSelection"
    Resume Fin
End Sub

Friend Sub RaiseOnVulnExpired(obj As acObject)
    If Valid(obj) Then RaiseEvent OnVulnExpired(obj)
End Sub
