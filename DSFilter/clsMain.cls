VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Decal Filter Interface
Implements INetworkFilter2

'Events
Public Event OnFilterDebugMessage(ByVal sMsg As String)
Public Event OnFilterErrorMessage(ByVal sMsg As String)

Public Event OnReady()
Public Event OnActionFailure(ByVal iReason As Integer)
Public Event OnMyDeath(ByVal sDeathMessage As String)
Public Event MeleeAttackComplete()
Public Event OnPlayerEvade(ByVal sAttacker As String)
Public Event OnTargetEvade(ByVal sTarget As String)
Public Event OnInflictMeleeDamage(ByVal sTarget As String, ByVal lDamage As Long, ByVal bCrit As Boolean, ByVal sDamageType As String)
Public Event OnReceiveMeleeDamage(ByVal sAttacker As String, ByVal lDamage As Long, ByVal bCrit As Boolean, ByVal sDamageType As String, ByVal sLocation As String)
Public Event OnReceiveDeathMessage(ByVal sMessage As String)
Public Event OnChangePlayerType(ByVal objPlayer As acObject, ByVal iOldType As ePlayerType)
Public Event OnReceiveTell(ByVal sMessage As String, ByVal sSenderName As String, ByVal lSenderGUID As Long, ByVal vData As Variant)
Public Event OnLocalChat(ByVal sMessage As String, ByVal sSenderName As String, ByVal lSenderID, ByVal lChatType As Long)
Public Event OnServerChat(ByVal sMessage As String, ByVal lType As Long)
Public Event OnFellowshipMessage(ByVal sSenderName As String, ByVal sMsg As String)
Public Event OnPlayerKill(ByVal objKiller As acObject, ByVal objKilled As acObject)

Public Event OnSpellCast(ByVal spellID As Long)
Public Event OnSpellCastOn(ByVal spellID As Long, ByVal objTarget As acObject)
Public Event OnSpellVisible(ByVal aName As String)

Public Event OnAddEnchantment(ByVal iSpellID As Long, ByVal lFamily As Long, ByVal dStartTime As Double, ByVal dTimeElapsed As Double, ByVal dDuration As Double, ByVal lSourceGUID As Long, ByVal iLayer As Long)
Public Event OnRemoveEnchantment(ByVal spellID As Long, ByVal iLayer As Long)
Public Event OnRemoveMultEnchantment(ByVal spellID As Long, ByVal iLayer As Long)
Public Event OnSilentRemoveEnchantment(ByVal spellID As Long, ByVal iLayer As Long)
Public Event OnSilentRemoveMultEnchantment(ByVal spellID As Long, ByVal iLayer As Long)

'Public Event OnObjectMoved(ByVal obj As acObject)

Public Event ForcedExit()

'Not working yet
Public Event OnPlayerAttack(objAttacker As acObject, objTarget As acObject)
Public Event OnAttackAnimation(objAttacker As acObject, objTarget As acObject)
'---

Public Event OnLogin()
Public Event OnLoginComplete()
Public Event OnEnterPortalSpace()
Public Event OnExitPortalSpace()

'Private Members
Private m_bLoaded As Boolean
Private m_Log As DebugLog
Private m_Account As New acAccount
Private m_Allegiance As acAllegiance
Private m_XpTracker As XpTrackerCls

Private m_locAccurate As acLoc      'Accurate Player position, taken from client side hook

Private m_bExitLoginPortal As Boolean
Private m_bInPortalSpace As Boolean
Private m_iServerId As eGameServer

Private WithEvents m_acHooks As Decal.ACHooks
Attribute m_acHooks.VB_VarHelpID = -1
Private WithEvents m_GameObjects As GameObjects
Attribute m_GameObjects.VB_VarHelpID = -1

'Clock
Private Const CLOCK_TIMER_INTERVAL = 100 'msec
Private Const CHECK_TIMER_INTERVAL = 30000 'msec -- 30 seconds
Private WithEvents m_tmrClock As Timer
Attribute m_tmrClock.VB_VarHelpID = -1
Public WithEvents m_tmrCheck As Timer
Attribute m_tmrCheck.VB_VarHelpID = -1
Private m_fLastTime As Double
Private m_fLastCheckTime As Double


'################################################################################
'                           CONSTRUCTOR / DESTRUCTOR
'################################################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    ' Persistant Data
    m_bLoaded = False
    
    'Set our reference
    Set g_Filter = Me
    
    'Create our log files
    Set m_Log = New DebugLog
        
    'Server won't change
    m_iServerId = SV_NONE
    
    'Clock
    Set m_tmrClock = frmTimers.tmrClock
    m_fLastTime = 0
    g_Time = 0
    m_tmrClock.Interval = CLOCK_TIMER_INTERVAL
    m_tmrClock.Enabled = False
    
    'Check Timer
    Set m_tmrCheck = frmTimers.tmrCheck
    m_fLastCheckTime = 0
    m_tmrCheck.Interval = CHECK_TIMER_INTERVAL
    m_tmrCheck.Enabled = False

    'Volatile Data
    Call ResetVolatileData
    
Fin:
    Exit Sub
ErrorHandler:
    DS_ErrorMsgBox "Class_Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub Class_Terminate()
    'Call ResetVolatileData
    'Call cleanCollections
    
    m_tmrClock.Enabled = False
    m_tmrCheck.Enabled = False

    Set m_tmrClock = Nothing
    Set m_tmrCheck = Nothing
    Set g_Filter = Nothing
    
    Set m_GameObjects = Nothing
    Set m_Allegiance = Nothing
    Set m_XpTracker = Nothing
    Set m_locAccurate = Nothing
    Set m_Account = Nothing
    
    Set g_Objects = Nothing
    Set g_Const = Nothing
    Set g_debugLog = Nothing
    Set g_errorLog = Nothing
    Set m_Log = Nothing
    
    g_Time = 0
    
    'Set m_acHooks = Nothing
    
End Sub

Private Sub ResetVolatileData()
On Error GoTo ErrorHandler

    'Game objects collections
    Set m_GameObjects = New GameObjects
    Set m_Allegiance = New acAllegiance
    Set m_XpTracker = New XpTrackerCls
    Set m_locAccurate = New acLoc
    
    Set g_Objects = m_GameObjects
    
    m_bExitLoginPortal = False
    m_bInPortalSpace = True
    
Fin:
    Exit Sub
ErrorHandler:
    DS_ErrorMsgBox "ResetVolatileData - " & Err.Description
    Resume Fin
End Sub

Private Sub cleanCollections()
On Error GoTo ErrorHandler

    Dim removeCol As colObjects
    Dim objEntity As acObject

    myDebug "Filter.cleanCollections called"
    myDebug "Filter.cleanCollections: g_Objects.Monsters: " & g_Objects.Monsters.Count
    myDebug "Filter.cleanCollections: g_Objects.Items: " & g_Objects.Items.World.Count
    
    Set removeCol = New colObjects
    
    ' First go thru the monster DB and remove any object marked dead
    For Each objEntity In g_Objects.Monsters
        If Valid(objEntity) Then
            ' build up list of objects to be removed
            If ((objEntity.Dead) Or (objEntity.canDelete)) And ((objEntity.timeData <> 0) And (g_Time > objEntity.timeData)) Then
                ' Remove Dead or objects that have been marked canDelete
                Call removeCol.AddObject(objEntity)
            End If
            
            'If (objEntity.timeData <> 0) And (objEntity.timeData > g_Time) Then
            '    'if timeData has passed, mark it for deletion
            '    objEntity.canDelete = True
            'End If
        End If
    Next
    
    For Each objEntity In removeCol
        'RemoveObject
        myDebug "Filter.cleanCollections: removing: " & objEntity.Name & " : " & objEntity.GUID
        Call m_GameObjects.RemoveObject(objEntity.GUID, "cleanCollections")
    Next
    
Fin:
    Set objEntity = Nothing
    Set removeCol = Nothing
    Exit Sub
ErrorHandler:
    myError "cleanCollections - " & Err.Description
    Resume Fin
End Sub

'################################################################################
'                              INTERFACE IMPL
'################################################################################

Private Sub INetworkFilter2_Initialize(ByVal pService As DecalNet.INetService)
On Error GoTo ErrorHandler

    m_bLoaded = False
    
    myDebug "Initializing Filter..."
    myDebug "App path : " & App.Path
    myDebug "Data Folder: " & GetDataFolder
    
    Set m_acHooks = pService.Hooks
    
    'Some global objects...
    Set g_Const = New acConst
    
    'Load the Filter data
    If Not LoadFilterData Then
        myError "Unable to load filter datas. Filter initialization failed."
        GoTo Fin
    End If
    
    'Init Clock
    m_fLastTime = timeGetTime
    g_Time = 0
    m_tmrClock.Enabled = True
    
    'Init Checker
    m_tmrCheck.Enabled = True

    'Loaded fine
    m_bLoaded = True
    
Fin:
    Exit Sub
ErrorHandler:
    m_bLoaded = False
    myError "INetworkFilter2_Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub INetworkFilter2_Terminate()

    m_tmrClock.Enabled = False
    m_tmrCheck.Enabled = False
    m_bLoaded = False
    
    Set m_acHooks = Nothing
    Set m_XpTracker = Nothing
    Set m_Account = Nothing
    Set m_Allegiance = Nothing
    Set m_GameObjects = Nothing
    
    Set g_Objects = Nothing
    Set g_Const = Nothing
    Set g_Filter = Nothing
    
End Sub

Private Sub m_tmrCheck_Timer()
On Error GoTo ErrorHandler
    
    'Every 30 seconds, clean out the old data
    Call cleanCollections
    
Fin:
    Exit Sub
ErrorHandler:
    myError "Filter.tmrCheck_Timer - " & Err.Description & " - Line: " & Erl
    Resume Fin
End Sub

Private Sub m_tmrClock_Timer()
On Error GoTo ErrorHandler
    
    Dim fDelta As Double
    Dim fCurTime As Double
    
    fCurTime = timeGetTime
    fDelta = fCurTime - m_fLastTime
    
    If (fDelta >= 1000) Then
        'Give current time in seconds
        m_fLastTime = fCurTime
        g_Time = g_Time + 1
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    myError "Filter.tmrClock_Timer - " & Err.Description & " - Line: " & Erl
    Resume Fin
End Sub

'################################################################################
'                               PROPERTIES
'################################################################################

Public Property Get Account() As acAccount
Attribute Account.VB_Description = "Account Info : characters list"
    Set Account = m_Account
End Property

Public Property Get Player() As acObject
Attribute Player.VB_Description = "GUID of the character logged in"
    If Valid(m_GameObjects) Then
        Set Player = m_GameObjects.Player
    End If
End Property

Public Property Get GameObjects() As GameObjects
    Set GameObjects = m_GameObjects
End Property

Public Property Get ds_ACHooks() As Decal.ACHooks
    Set ds_ACHooks = m_acHooks
End Property

Public Property Get LoginComplete() As Boolean
    LoginComplete = m_bExitLoginPortal
End Property

Public Property Get InPortalSpace() As Boolean
    InPortalSpace = m_bInPortalSpace
End Property

Public Property Get Allegiance() As acAllegiance
    Set Allegiance = m_Allegiance
End Property

Public Property Get Loaded() As Boolean
    Loaded = m_bLoaded
End Property

Public Property Get ACConstants() As acConst
    Set ACConstants = g_Const
End Property

Public Property Get XpTracker() As XpTrackerCls
    Set XpTracker = m_XpTracker
End Property

Public Property Get AccuratePlayerLoc() As acLoc
On Error GoTo ErrorHandler

    With m_locAccurate
        .Landblock = m_acHooks.Landcell
        .Xoff = m_acHooks.LocationX
        .Yoff = m_acHooks.LocationY
        .Zoff = m_acHooks.LocationZ
    End With
    
Fin:
    Set AccuratePlayerLoc = m_locAccurate
    Exit Property
ErrorHandler:
    myError "AccuratePlayerLoc - " & Err.Description
    Resume Fin
End Property

Public Property Get ServerID() As eGameServer
    ServerID = m_iServerId
End Property

Public Property Get ServerName() As String
On Error GoTo ErrorHandler

    ServerName = g_Const.GetServerName(m_iServerId)

Fin:
    Exit Property
ErrorHandler:
    ServerName = "Unknown Server"
    myError "ServerName - " & Err.Description
    Resume Fin
End Property

Public Property Get Time() As Long
    Time = g_Time
End Property

'################################################################################
'                          Public Methods
'################################################################################

Public Sub SetDebugMode(ByVal aVal As Boolean)
    m_Log.bDebugMode = aVal
    myDebug "SetDebugMode: " & aVal
End Sub

'################################################################################
'                             FRIEND METHODS
'################################################################################

Friend Sub LogDebugMsg(ByVal sMsg As String)
On Error Resume Next
    Call m_Log.DebugMsg(sMsg)
    RaiseEvent OnFilterDebugMessage(sMsg)
End Sub

Friend Sub LogErrorMsg(ByVal sMsg As String)
On Error Resume Next
    Call m_Log.ErrorMsg(sMsg)
    RaiseEvent OnFilterErrorMessage(sMsg)
End Sub

'Files loading
Private Function LoadFilterData() As Boolean
On Error GoTo ErrorHandler
    
    If Not g_Const.LoadMaterialsList(GetDataFolder & "\" & FOLDER_DATA & "\" & FILE_MATERIALS) Then
        myError "LoadFilterData - Failed to load Materials List"
        GoTo Fin
    End If
    
    LoadFilterData = True
    
Fin:
    Exit Function
ErrorHandler:
    LoadFilterData = False
    myError "LoadFilterData - " & Err.Description
    Resume Fin
End Function

'################################################################################
'                              ACHOOKS EVENTS
'################################################################################

Private Sub m_acHooks_OnSelectItem(ByVal lGUID As Long)
On Error GoTo ErrorHandler

    Call m_GameObjects.SetCurrentSelection(lGUID)

Fin:
    Exit Sub
ErrorHandler:
    myError "m_acHooks_OnSelectItem"
    Resume Fin
End Sub

Private Sub m_acHooks_StatusTextIntercept(ByVal bstrText As String, bEat As Boolean)
On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = LCase(bstrText)

    If InStr(sMsg, "too busy") Then
        RaiseEvent OnActionFailure(FAIL_BUSY)

    ElseIf InStr(1, sMsg, "not ready") Then
        RaiseEvent OnActionFailure(FAIL_NOT_READY)

    ElseIf InStr(1, sMsg, "you can only move one item") Then    'can only move one item at a time
        RaiseEvent OnActionFailure(FAIL_CAN_ONLY_MOVE_ONE)

    ElseIf InStr(1, sMsg, "is already in use") Then
        RaiseEvent OnActionFailure(FAIL_ALREADY_IN_USE)

    ElseIf InStr(1, sMsg, "unable to move to object") Then
        RaiseEvent OnActionFailure(FAIL_UNABLE_TO_MOVE_TO_OBJECT)

    ElseIf InStr(1, sMsg, "cannot be used") Then
        RaiseEvent OnActionFailure(FAIL_CANNOT_BE_USED)

    End If

Fin:
    Exit Sub
ErrorHandler:
    myError "m_achooks_OnToolText - " & Err.Description
    Resume Fin
End Sub

Private Sub m_acHooks_ObjectDestroyed(ByVal lGUID As Long)
On Error GoTo ErrorHandler

    myDebug "acHooks.ObjectDestroyed called: " & lGUID
    If Valid(g_Objects.FindObject(lGUID)) Then
        'myError "WARNING: acHooks.ObjectDestroyed: FOUND VALID OBJECT: " & g_Objects.FindObject(lGUID).Name & " : " & lGUID
        'Call m_GameObjects.RemoveObject(lGUID, "acHooks.ObjectDestroyed")
    End If

Fin:
    Exit Sub
ErrorHandler:
    myError "m_achooks_ObjectDestroyed(" & lGUID & ") - " & Err.Description
    Resume Fin
End Sub



'################################################################################
'                              NET EVENTS
'################################################################################

'Message from Client -> Server
Private Sub INetworkFilter2_DispatchClient(ByVal Message As DecalNet.IMessage2)
On Error GoTo ErrorHandler
    
    If Not m_bLoaded Then
        Exit Sub
    End If

    'myDebug "DispatchClient: " & Message.Type

    Select Case Message.Type
        '---------------- Outbound Game Events --------------------
        Case MSG_GAME_ACTION
            Select Case Message.Value("action")
                Case ACT_CAST_SPELL
                    Call NetAction_CastSpell(Message)
                Case ACT_CAST_SPELL_ON
                    Call NetAction_CastSpellOn(Message)
            End Select
    
    End Select

Fin:
    Exit Sub
ErrorHandler:
    myError "INetworkFilter2_DispatchClient: " & Err.Description
    Resume Fin
End Sub


'Message from Server -> Client
Private Sub INetworkFilter2_DispatchServer(ByVal Message As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    If Not m_bLoaded Then
        Exit Sub
    End If

    Select Case Message.Type
    
        Case MSG_SET_CHAR_QWORD
            ' Total/Free XP
            'myError "- Set Char QWord -"
            Call m_XpTracker.UpdateXp(Message)
            
        Case MSG_SET_CHAR_DWORD
            'myDebug "- Set Char DWord -"
            Call m_XpTracker.UpdateXp(Message)

        Case MSG_SERVER_NAME
            'myDebug "- Server Name -"
            Call NetEvent_ServerName(Message)
            
        Case MSG_CHAR_LIST
            Set m_Account = New acAccount
            Call Net_CharacterList(Message)
            
        'Case RETIRED_MSG_START_3DMODE
        '    myDebug "- Start 3D Mode -"
            'Call ResetVolatileData
        
        Case MSG_END_3DMODE
            myDebug "- End 3D Mode -"
            Call ResetVolatileData
            RaiseEvent ForcedExit
        
        '---------------------------------------------------
    
        Case MSG_LOGIN_CHAR
            Player.GUID = Message.Value("character")
            myDebug "[MSG_LOGIN_CHAR] Character GUID : " & Player.GUID
        
        Case MSG_TOGGLE_VISIBILITY
            'myDebug "- Toggle Visibility -"
            Call NetEvent_OnToggleVisibility(Message)
            
        Case MSG_MSG, MSG_MSG_RANGED
            'myDebug "- Msg Ranged -"
            Call NetEvent_OnReceiveLocalChat(Message)
        
        Case MSG_SERVER_CHAT
            'myDebug "- Server Chat -"
            Call NetEvent_OnReceiveServerChat(Message)

        Case MSG_TURBINE_CHAT
            'myDebug "- Turbine Chat -"
            Call NetEvent_OnReceiveTurbineChat(Message)
            
        'Case RETIRED_MSG_SET_COVERAGE
            'Call NetEvent_OnSetCoverage(Message)

         '---------------- Object Events --------------------
         
        Case MSG_CREATE_OBJECT
            'myDebug "- CreateObject -"
            Call m_GameObjects.NetEvent_OnCreateObject(Message, False)
            
        Case MSG_UPDATE_OBJECT
            'myDebug "- UpdateObject -"
            Call m_GameObjects.NetEvent_OnCreateObject(Message, True)
        
        Case MSG_SET_POSITION
            'myDebug "- SetPosition -"
            Call m_GameObjects.NetEvent_OnSetPosition(Message)
            
        'Case MSG_SET_CHAR_POSITION
            'myDebug "- SetCharacterPosition -"
            'Call m_GameObjects.NetEvent_OnSetCharPosition(Message)
                
        Case MSG_DESTROY_OBJECT
            'myDebug "- Destroy Object -"
            Call m_GameObjects.RemoveObject(Message.Value("object"), "MSG_DESTROY_OBJECT")
        
        Case MSG_REMOVE_OBJECT
            'myDebug "- RemoveObject -"
            Call m_GameObjects.RemoveObject(Message.Value("object"), "MSG_REMOVE_OBJECT")
            
        Case MSG_ADJUST_STACK_SIZE
            'myDebug "- AdjustStackSize -"
            Call m_GameObjects.Items.NetEvent_OnAdjustStackSize(Message)
            
        'Case RETIRED_MSG_SET_WIELDER_CONTAINER
            'Call m_GameObjects.Items.NetEvent_OnSetWielderContainer(Message)
                    
        Case MSG_MOVE_INVENTORY_OBJECT
            'myDebug "- Move Inventory Object -"
            Call m_GameObjects.Items.NetEvent_OnMoveObjectIntoInventory(Message)
            
        Case MSG_WIELD_OBJECT
            'myDebug "- Wield Object -"
            Call m_GameObjects.Items.NetEvent_OnWieldObject(Message)
            
        Case MSG_APPLY_VISUAL_EFFECTS
            'myDebug "- Apply Visual Effects -"
            Call NetEvent_OnApplyVisualEffects(Message)
                
        'Case MSG_APPLY_SOUND_EFFECT
        '    Call NetEvent_OnApplySoundEffects(Message)
            
        Case MSG_ANIMATION
            'myDebug "- Animation -"
            Call NetEvent_OnAnimation(Message)
                
        Case MSG_PLAYER_KILL
            'myDebug "- Player Kill -"
            Call NetEvent_OnPlayerKill(Message)
                        
        'Case RETIRED_MSG_ATTACK
            'Call NetEvent_OnAttack(Message)
                
        '---------------- Inbound Game Events --------------------
        Case MSG_GAME_EVENT
            'myDebug "- Msg Game Event -"
            Select Case Message.Value("event")
                Case EV_LOGIN
                    Call NetEvent_OnLogin(Message)
                    
                Case EV_ALLEGIANCE_INFO
                    Call m_Allegiance.NetEvent_OnAllegianceInfo(Message)
                    
                Case EV_ACTION_FAILURE
                    RaiseEvent OnActionFailure(Message.Value("type"))
                    
                Case EV_READY
                    RaiseEvent OnReady
                
                Case EV_MY_DEATH
                    RaiseEvent OnMyDeath(Message.Value("text"))
                       
                Case EV_MELEE_ATTACK_COMPLETE
                    RaiseEvent MeleeAttackComplete
                
                Case EV_INFLICT_MELEE_DMG
                    Call NetEvent_OnInflictMeleeDamage(Message)
                    
                Case EV_RECEIVE_MELEE_DMG
                    Call NetEvent_OnReceiveMeleeDamage(Message)
                
                Case EV_TARGET_EVADES_ATTACK
                    Call NetEvent_OnTargetEvade(Message)
                    
                Case EV_PLAYER_EVADES_ATTACK
                    Call NetEvent_OnPlayerEvade(Message)
                
                Case EV_DEATH_MESSAGE
                    RaiseEvent OnReceiveDeathMessage(Message.Value("text"))
             
                '------- Chat Events --------
                
                Case EV_DIRECT_CHAT 'receiving @tell
                    Call NetEvent_OnDirectChat(Message)
                    
                Case EV_GROUP_CHAT
                    Call NetEvent_OnGroupChat(Message)
                
                '---------------- Items/Inventory Management --------------------
                
                Case EV_IDENTIFY_OBJECT
                    Call m_GameObjects.NetEvent_OnIdentifyObject(Message)
                 
                Case EV_WEAR_ITEM 'Wear/Equip item
                    Call m_GameObjects.Items.NetEvent_OnWearItem(Message)
                    
                Case EV_INSERT_INVENTORY_ITEM 'Instert Item in inventory
                    Call m_GameObjects.Items.NetEvent_OnInsertInventoryItem(Message)
                    
                Case EV_DROP_ITEM
                    Call m_GameObjects.Items.NetEvent_OnDropFromInventory(Message)
                    
                Case EV_SET_PACK_CONTENT
                    Call m_GameObjects.Items.NetEvent_OnSetPackContents(Message)
                
                Case EV_APPROACH_VENDOR
                    Call m_GameObjects.NetEvent_OnApproachVendor(Message)

                '---------------- Fellowship events --------------------
                
                Case EV_FELLOWSHIP_CREATE
                    Call g_Objects.Fellowship.HandleFellowshipEvents(Message, EV_FELLOWSHIP_CREATE)
                    
                Case EV_FELLOWSHIP_QUIT
                    Call g_Objects.Fellowship.HandleFellowshipEvents(Message, EV_FELLOWSHIP_QUIT)
                    
                Case EV_FELLOWSHIP_DISBANDS
                    Call g_Objects.Fellowship.HandleFellowshipEvents(Message, EV_FELLOWSHIP_DISBANDS)
                    
                Case EV_FELLOWSHIP_RECRUIT
                    Call g_Objects.Fellowship.HandleFellowshipEvents(Message, EV_FELLOWSHIP_RECRUIT)
                    
                Case EV_FELLOWSHIP_DISMISS
                    Call g_Objects.Fellowship.HandleFellowshipEvents(Message, EV_FELLOWSHIP_DISMISS)
                    
                Case RETIRED_EV_FELLOWSHIP_INVITATION
                    Call g_Objects.Fellowship.HandleFellowshipEvents(Message, RETIRED_EV_FELLOWSHIP_INVITATION)
                
                '----------- Spell Events -------------
                Case EV_ADD_ENCHANTMENT
                    Call NetEvent_OnAddEnchantment(Message)

                Case EV_REMOVE_ENCHANTMENT
                    Call NetEvent_OnRemoveEnchantment(Message)
                
                Case EV_REMOVE_MULT_ENCHANTMENT
                    Call NetEvent_OnRemoveMultEnchantment(Message)
                
                Case EV_REMOVE_ENCHANTMENT_SILENT
                    Call NetEvent_OnSilentRemoveEnchantment(Message)
                
                Case EV_REMOVE_MULT_ENCHANTMENT_SILENT
                    Call NetEvent_OnSilentRemoveMultEnchantment(Message)
                
                '----------- Unknown Game Event -------------
                Case Else
                    myDebug "MSG_GAME_EVENT: unknown: Event: " & Message.Value("event")

                End Select
                
                    
        '---------------- Other Messages --------------------

    End Select
    
Fin:
    'myDebug "INetworkFilter2_DispatchServer Done"
    Exit Sub
ErrorHandler:
    myError "INetworkFilter2_DispatchServer - " & Err.Description
    Resume Fin
End Sub


Private Sub NetEvent_ServerName(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sMsg As String

    sMsg = UCase(pMsg.Value("server"))

    myDebug "NetEvent_ServerName: " & sMsg
    
    If InStr(sMsg, "WINTERSEBB") Then
        m_iServerId = SV_WINTERSEBB
        
    ElseIf InStr(sMsg, "SOLCLAIM") Then
        m_iServerId = SV_SOLCLAIM
        
    ElseIf InStr(sMsg, "MORNINGTHAW") Then
        m_iServerId = SV_MORNINGTHAW
        
    ElseIf InStr(sMsg, "HARVESTGAIN") Then
        m_iServerId = SV_HARVESTGAIN
        
    ElseIf InStr(sMsg, "FROSTFELL") Then
        m_iServerId = SV_FROSTFELL
        
    ElseIf InStr(sMsg, "THISTLEDOWN") Then
        m_iServerId = SV_THISTLEDOWN
        
    ElseIf InStr(sMsg, "LEAFCULL") Then
        m_iServerId = SV_LEAFCULL
        
    ElseIf InStr(sMsg, "DARKTIDE") Then
        m_iServerId = SV_DARKTIDE
        
    ElseIf InStr(sMsg, "VERDANTINE") Then
        m_iServerId = SV_VERDANTINE
        
    Else
        myError "NetEvent_ServerName: couldn't find Server ID"
        m_iServerId = SV_NONE
    End If
    
    myDebug "ServerID : " & m_iServerId
    
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_MOTD - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnLogin(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    myDebug "NetEvent_OnLogin"
    
    'Catch and set Total pyreals
    RaiseEvent OnLogin
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnLogin - " & Err.Description
    Resume Fin
End Sub


'
Private Sub NetEvent_OnPlayerEvade(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sAttacker As String
    sAttacker = pMsg.Value("attacker")
    
    RaiseEvent OnPlayerEvade(sAttacker)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnPlayerEvade - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnTargetEvade(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sTarget As String
    sTarget = pMsg.Value("target")
    
    RaiseEvent OnTargetEvade(sTarget)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnTargetEvade - " & Err.Description
    Resume Fin
End Sub

'NetEvent_OnInflictMeleeDamage
Private Sub NetEvent_OnInflictMeleeDamage(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sTarget As String
    Dim lDamageType As Long
    Dim lDamage As Long
    Dim bCrit As Boolean
    
    Dim sDamageType As String
    
    sTarget = pMsg.Value("target")
    lDamageType = pMsg.Value("type")
    lDamage = pMsg.Value("damage")
    bCrit = (pMsg.Value("critical") = 1)
    
    
    'DamageType
    '0x01 Slashing
    '0x02 Piercing
    '0x04 Bludgeoning
    '0x08 Cold
    '0x10 Fire
    '0x20 Acid
    '0x40 Electric
    
    Select Case lDamageType
        Case &H1&
            sDamageType = "slash"
        Case &H2&
            sDamageType = "pierce"
        Case &H4&
            sDamageType = "bludge"
        Case &H8&
            sDamageType = "cold"
        Case &H10&
            sDamageType = "fire"
        Case &H20&
            sDamageType = "acid"
        Case &H40&
            sDamageType = "electric"
    End Select
    
    RaiseEvent OnInflictMeleeDamage(sTarget, lDamage, bCrit, sDamageType)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnInflictMeleeDamage - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnReceiveMeleeDamage(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim objAttacker As acObject
    
    Dim sAttacker As String
    Dim lDamageType As Long
    Dim lLocation As Long
    Dim lDamage As Long
    Dim bCrit As Boolean
    
    Dim sDamageType As String
    Dim sLocation As String
    
1    sAttacker = pMsg.Value("attacker")
2    lDamageType = pMsg.Value("type")
3    lLocation = pMsg.Value("location")
4    lDamage = pMsg.Value("damage")
5    bCrit = (pMsg.Value("critical") = 1)
    
6    'Set objAttacker = g_Objects.FindObject(sAttacker, False)
7    'If Not Valid(objAttacker) Then
8    '    myDebug "NetEvent_OnDamageReceived - Attacker -- " & sAttacker & "  -- is NOT A VALID OBJECT!"
9    'End If
    
    'DamageType
    '0x01 Slashing
    '0x02 Piercing
    '0x04 Bludgeoning
    '0x08 Cold
    '0x10 Fire
    '0x20 Acid
    '0x40 Electric
    
    Select Case lDamageType
        Case &H1&
            sDamageType = "slash"
        Case &H2&
            sDamageType = "pierce"
        Case &H4&
            sDamageType = "bludge"
        Case &H8&
            sDamageType = "cold"
        Case &H10&
            sDamageType = "fire"
        Case &H20&
            sDamageType = "acid"
        Case &H40&
            sDamageType = "electric"
    End Select
    
    'The DamageLocation indicates where damage was done.
    '0x00 Head
    '0x01 Chest
    '0x02 Abdomen
    '0x03 Upper Arm
    '0x04 Lower Arm
    '0x05 Hand
    '0x06 Upper Leg
    '0x07 Lower Leg
    '0x08 Foot
    Select Case lLocation
        Case &H0&
            sLocation = "head"
        Case &H1&
            sLocation = "chest"
        Case &H2&
            sLocation = "abdomen"
        Case &H3&
            sLocation = "upper arm"
        Case &H4&
            sLocation = "lower arm"
        Case &H5&
            sLocation = "hand"
        Case &H6&
            sLocation = "upper leg"
        Case &H7&
            sLocation = "lower leg"
        Case &H8&
            sLocation = "foot"
    End Select
    
    RaiseEvent OnReceiveMeleeDamage(sAttacker, lDamage, bCrit, sDamageType, sLocation)
    
Fin:
    Set objAttacker = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnReceiveMeleeDamage - " & Err.Description & " line: " & Erl
    Resume Fin
End Sub

Private Sub NetEvent_OnAttack(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim objAttacker As acObject
    Dim objTarget As acObject
    
    myDebug "NetEvent_OnAttack - Player #" & pMsg.Value("attacker") & " is attacking #" & pMsg.Value("target")
        
    Set objAttacker = g_Objects.FindObject(pMsg.Value("attacker"), False)
    Set objTarget = g_Objects.FindObject(pMsg.Value("target"), False)
    
    If Valid(objAttacker) And Valid(objTarget) Then
        myDebug "NetEvent_OnAttack - Player " & objAttacker.Name & " is attacking " & objTarget.Name
        RaiseEvent OnPlayerAttack(objAttacker, objTarget)
    End If
    
Fin:
    Set objAttacker = Nothing
    Set objTarget = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnAttack - " & Err.Description
    Resume Fin
End Sub

'@tell
Private Sub NetEvent_OnDirectChat(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler
    
    Dim sMessage As String
    Dim sName As String
    Dim lSenderID As Long
    Dim lChatType As Long

    sName = pMsg.Value("senderName")
    lSenderID = pMsg.Value("sender")
    sMessage = pMsg.Value("text")
    lChatType = pMsg.Value("type")
    
'0x00    Broadcast (e.g. allegiance MOTD)
'0x02    Public Chat
'0x03    Private Tell
'0x04    Outgoing Tell (e.g. 'You tell ...')
'0x07    Magic Spell Results
'0x0c    NPC Chat
'0x11    Player Spellcasting
'0x12    Creature Chat (e.g. 'Fellow warriors, aid me!')
'0x17    Recall (e.g. 'Player is recalling home.')
      
    If (lChatType And &H3&) Then
        RaiseEvent OnReceiveTell(sMessage, sName, lSenderID, pMsg.RawData)
    End If

    myDebug "NetEvent_OnDirectChat (" & lChatType & ") - " & sName & " (" & lSenderID & ") : " & sMessage

Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnGroupChat - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnGroupChat(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler
    
    Dim sMessage As String
    Dim sName As String
    Dim lChatType As Long
    Dim lSenderID As Long

    sName = pMsg.Value("senderName")
    lSenderID = pMsg.Value("sender")
    sMessage = pMsg.Value("text")
    lChatType = pMsg.Value("type")
    
    If sName = "" Then
        sName = g_Objects.Player.Name
    End If
    
    If (lChatType And MSK_FOLLOW_TO_MONARCH) Then
        If (m_GameObjects.Player.MonarchID = m_GameObjects.Player.GUID) Then      'Character is a monarch.
            RaiseEvent OnReceiveTell(sMessage, sName, lSenderID, pMsg.RawData)
        End If
    End If
    
    If (lChatType And MSK_FELLOW) Then
        RaiseEvent OnFellowshipMessage(sName, sMessage)
    ElseIf (lChatType And MSK_ALLEGIANCE) Then
        ' allegiance warning
    End If
    
    myDebug "NetEvent_OnGroupChat (" & lChatType & ") - " & sName & " : " & sMessage
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnGroupChat - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnReceiveLocalChat(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sMessage As String
    Dim sName As String
    Dim lChatType As Long
    Dim lSenderID As Long

    sName = pMsg.Value("senderName")
    lSenderID = pMsg.Value("sender")
    sMessage = pMsg.Value("text")
    lChatType = pMsg.Value("type")

    myDebug "NetEvent_OnReceiveLocalChat: name: " & sName & ":" & lChatType & " (" & lSenderID & ") : " & sMessage

'0x00    Broadcast (e.g. allegiance MOTD)
'0x02    Public Chat
'0x03    Private Tell
'0x04    Outgoing Tell (e.g. 'You tell ...')
'0x07    Magic Spell Results
'0x0c    NPC Chat
'0x11    Player Spellcasting
'0x12    Creature Chat (e.g. 'Fellow warriors, aid me!')
'0x17    Recall (e.g. 'Player is recalling home.')

    RaiseEvent OnLocalChat(sMessage, sName, lSenderID, lChatType)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnReceiveLocalChat - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnReceiveTurbineChat(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sMessage As String
    Dim sSenderName As String
    Dim lSenderGUID As Long
    Dim lChatType As Long
    Dim lChannel As Long

    'sMessage = pMsg.Value("text")
    'sSenderName = pMsg.Value("senderName")
    'lSenderGUID = pMsg.Value("sender")
    
    ' 0x01 - Inbound
    ' 0x03 - Outbound
    lChatType = pMsg.Value("type")
    
    If (lChatType And &H1&) Then
        sMessage = pMsg.Value("text")
        sSenderName = pMsg.Value("senderName")
        lSenderGUID = pMsg.Value("sender")
        
        myDebug "NetEvent_OnReceiveTurbineChat: " & lChatType & ":" & sSenderName & " (" & lSenderGUID & ") -- " & sMessage
        
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnReceiveTurbineChat - " & Err.Description
    Resume Fin
End Sub


Private Sub NetEvent_OnReceiveServerChat(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim sMessage As String
    Dim lChatType As Long

    sMessage = pMsg.Value("text")
    lChatType = pMsg.Value("type")

    myDebug "NetEvent_OnReceiveServerChat: " & lChatType & " : " & sMessage

'0x00    Broadcast (e.g. allegiance MOTD)
'0x02    Public Chat
'0x03    Private Tell
'0x04    Outgoing Tell (e.g. 'You tell ...')
'0x07    Magic Spell Results
'0x0c    NPC Chat
'0x11    Player Spellcasting
'0x12    Creature Chat (e.g. 'Fellow warriors, aid me!')
'0x17    Recall (e.g. 'Player is recalling home.')

    RaiseEvent OnServerChat(sMessage, lChatType)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnReceiveServerChat - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnSetCoverage(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim lObjGUID As Long
    Dim lFlag As Long
    Dim lCoverage As Long
    Dim objEntity As acObject
    
    lObjGUID = pMsg.Value("object")
    lFlag = pMsg.Value("unknown")
    lCoverage = pMsg.Value("coverage")
    
    Set objEntity = m_GameObjects.FindObject(lObjGUID)
    
    If Valid(objEntity) Then
        'myDebug "NetEvent_OnSetCoverage : " & objEntity.Name & " - Flag: " & lFlag & " - Coverage: " & lCoverage
    
        If objEntity.ObjectType = TYPE_PLAYER Then
            Dim oldPlayerType As ePlayerType
            oldPlayerType = objEntity.PlayerType
            
            If lFlag = &H86& Then   'PK state change message?
                Select Case lCoverage
                    Case &H40&  'PKL
                        objEntity.PlayerType = PLAYER_PINK
                        
                    Case &H4&   'PK
                        objEntity.PlayerType = PLAYER_RED
                        
                    Case &H2&
                        objEntity.PlayerType = PLAYER_WHITE
                End Select
                
                If objEntity.PlayerType <> oldPlayerType Then
                    RaiseEvent OnChangePlayerType(objEntity, oldPlayerType)
                End If
                
            End If
        End If
    End If

Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnSetCoverage - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnToggleVisibility(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler
    
    Dim lCharGUID As Long
    Dim iPortalType As Integer
    Dim pString As String
    
    lCharGUID = pMsg.Value("object")
    iPortalType = CInt(pMsg.Value("portalType"))
    
    'myDebug "OnToggleVisibility : " & g_Objects.FindObject(lCharGUID).Name & " (PlayerGUID: " & g_Objects.Player.GUID & ") - Portal Type : " & Hex(iPortalType)
    
    If (lCharGUID = Player.GUID) Then
        Select Case iPortalType
            Case PORTAL_TYPE_ENTER
                m_bInPortalSpace = True
                RaiseEvent OnEnterPortalSpace
                
            Case PORTAL_TYPE_EXIT
                m_bInPortalSpace = False
                If Not m_bExitLoginPortal Then  'First login?
                    m_bExitLoginPortal = True
                    RaiseEvent OnLoginComplete
                Else
                    RaiseEvent OnExitPortalSpace
                End If
        End Select
    End If
    
    pString = Hex(iPortalType)
    
    myDebug "OnToggleVisibility : " & g_Objects.FindObject(lCharGUID).Name & " - PType (Hex): " & Hex(iPortalType) & " pString: " & pString
    
    If InStr(pString, "8374") Or InStr(pString, "8774") Or InStr(pString, "374") Then
            myDebug "INSTR: raising OnSpellVisible: " & g_Objects.FindObject(lCharGUID).Name
            RaiseEvent OnSpellVisible(g_Objects.FindObject(lCharGUID).Name)
    End If
    
    'Select Case pString
    '    Case PS_BOLT Or PS_ARC Or PS_STREAK
    '        myDebug "PS: raising OnSpellVisible: " & g_Objects.FindObject(lCharGUID).Name
    '        RaiseEvent OnSpellVisible(g_Objects.FindObject(lCharGUID).Name)
    '    Case 8374 Or 8774 Or 374
    '        myDebug "H: raising OnSpellVisible: " & g_Objects.FindObject(lCharGUID).Name
    '        RaiseEvent OnSpellVisible(g_Objects.FindObject(lCharGUID).Name)
    'End Select
   
Fin:
    Exit Sub
ErrorHandler:
    myError "OnToggleVisibility - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnApplySoundEffects(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim iVuln As Integer
    Dim objEntity As acObject
    Dim lObjectGUID As Long
    Dim lEffectID As Long
    Dim dparameter As Double
    
    lObjectGUID = pMsg.Value("object")
    lEffectID = pMsg.Value("effect")
    dparameter = pMsg.Value("parameter")

    'Audio/Visual Effect ID
    '0x04    War Launch
    '0x05    War Land
    '0x2B    Red Shield Falling (Fire Vulnerability)
    '0x2D    Orange Shield Falling (Piercing Vulnerability)
    '0x2F    Yellow Shield Falling (Blade Vulnerability)
    '0x31    Green Shield Falling (Acid Vulnerability)
    '0x33    Cyan Shield Falling (Cold Vulnerability)
    '0x35    Purple Shield Falling (Lightning Vulnerability)
    '0x37    Black Shield Falling (Bludgeon Vulnerability, Imperil)
    

Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnApplySoundEffects - " & Err.Description
    Resume Fin
End Sub


Private Sub NetEvent_OnApplyVisualEffects(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim iVuln As Integer
    Dim objEntity As acObject
    Dim lObjectGUID As Long
    Dim lEffectID As Long
    
    lObjectGUID = pMsg.Value("object")
    lEffectID = pMsg.Value("effect")
    
    myDebug "OnApplyVisualEffect: Hex: " & Hex(lEffectID) & " on GUID: " & lObjectGUID
    myDebug "OnApplyVisualEffect: Raw: " & lEffectID & " on GUID: " & lObjectGUID
    
    If lEffectID = EF_WAR_LAUNCH Then
        RaiseEvent OnSpellVisible("WAR_LAUNCH Event")
    ElseIf lEffectID = EF_WAR_LAND Then
        RaiseEvent OnSpellVisible("WAR_LAND Event")
    ElseIf lEffectID = EF_FIZZLE Then
        RaiseEvent OnSpellVisible("FIZZLE Event")
    ElseIf lEffectID = EF_EQUIP_ITEM Then
    ElseIf lEffectID = EF_UNEQUIP_ITEM Then
    End If
    
    iVuln = -1
    Call m_GameObjects.Exists(lObjectGUID, objEntity)
    
    If Not Valid(objEntity) Then
        'Spells don't get added to GameObjects, so not an error
        'myError "NetEvent_OnApplyVisualEffects: No valid objEntity found in m_GameObjects GUID: " & lObjectGUID
        Set objEntity = New acObject
        myDebug "NetEvent_OnApplyVisualEffects: No valid objEntity found in m_GameObjects GUID: " & lObjectGUID
    End If
    
    Select Case lEffectID
        Case EF_FIRE_VULN
            iVuln = DMG_FIRE
        Case EF_PIERCE_VULN
            iVuln = DMG_PIERCING
        Case EF_BLADE_VULN
            iVuln = DMG_SLASHING
        Case EF_COLD_VULN
            iVuln = DMG_COLD
        Case EF_ACID_VULN
            iVuln = DMG_ACID
        Case EF_LIGHTNING_VULN
            iVuln = DMG_LIGHTNING
        Case EF_IMPERIL_BLUDG
            myDebug "OnApplyVisualEffect: Imp/Bludge landing on " & lObjectGUID
            objEntity.ImpOrBludgeon = True
            Call m_GameObjects.FireObjectImpOrBludgeoned(objEntity)
        Case EF_YIELD
            myDebug "OnApplyVisualEffect: Yield landing on " & lObjectGUID
            objEntity.Yielded = True
            Call m_GameObjects.FireOnObjectYielded(objEntity)
            
    End Select
    
    If iVuln <> -1 Then
        Call objEntity.AddVuln(iVuln)
        Call m_GameObjects.FireObjectVulned(objEntity, iVuln)
        myDebug "OnApplyVisualEffect: VULN landed: " & iVuln & " on target: " & objEntity.Name
    End If
    
Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnApplyVisualEffects - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnAnimation(pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim lObjGUID As Long
    Dim iAnimType As Integer
    Dim objEntity As acObject
    Dim bHasTarget As Boolean
    
    iAnimType = -1
    lObjGUID = pMsg.Value("object")
    iAnimType = pMsg.Value("animation_type")
    bHasTarget = (pMsg.Value("type_flags") = 1)
    
    Set objEntity = m_GameObjects.FindObject(lObjGUID, False)
    
    'If object not created yet, see if we can detect missing critters
    'If Not Valid(objEntity) And (iAnimType = &H6&) Then
    '    'Attack animation
    '    If bHasTarget Then
    '        Dim objTarget As acObject
    '
    '        Set objEntity = New acObject
    '        objEntity.GUID = lObjGUID
    '        objEntity.ObjectType = TYPE_MONSTER
    '
    '        'Set objTarget = g_Objects.FindObject(pMsg.Value("target"))
    '
    '        objEntity.Loc.Landblock = pMsg.Value("landblock")
    '        objEntity.Loc.Xoff = pMsg.Value("xOffset")
    '        objEntity.Loc.Yoff = pMsg.Value("yOffset")
    '        objEntity.Loc.Zoff = pMsg.Value("zOffset")
    '
    '        If Not g_Objects.Monsters.AddObject(objEntity) Then
    '            myError "GameObjects.NetEcho_OnAnimation: failed to add monster " & objEntity.GUID & " to Monsters collection"
    '        Else
    '            myDebug "GameObjects.NetEcho_OnAnimation: adding m_colMonsters: " & objEntity.GUID & " to Monsters collection"
    '        End If
    '
    '    End If
    'End If


    If iAnimType = 0 And Valid(objEntity) Then
        Dim lFlags As Long
        lFlags = pMsg.Value("flags")
        
        If lFlags And &H2 Then
            Dim iAnim1 As Integer
            iAnim1 = pMsg.Value("animation_1")
            
            If iAnim1 = 17 And (objEntity.ItemType <> ITEM_CORPSE) Then
                myDebug "Anim - iAnimType:" & iAnim1 & " - Atker:" & objEntity.Name
                objEntity.Dead = True
                Call m_GameObjects.FireOnObjectDeath(objEntity)
            End If
            
        End If
    End If
    
    'If Valid(objEntity) And (iAnimType = &H6&) And bHasTarget Then
    '    Dim lLandblock As Long
    '
    '    lLandblock = pMsg.Value("landblock")
    '
    '    If lLandblock <> 0 Then
    '        Dim oldLoc As acLoc
    '        Set oldLoc = objEntity.Loc.Clone
    '
    '        objEntity.Loc.Landblock = lLandblock
    '        objEntity.Loc.Xoff = pMsg.Value("xOffset")
    '        objEntity.Loc.Yoff = pMsg.Value("yOffset")
    '        objEntity.Loc.Zoff = pMsg.Value("zOffset")
    '
    '        If Not objEntity.Loc.Equals(oldLoc) Then
    '            RaiseEvent OnObjectMoved(objEntity)
    '            objEntity.timeData = 0
    '        End If
    '        Set oldLoc = Nothing
    '    End If
    'End If
    
'    'Attack animation
'    'If bHasTarget Then
'        Dim iStance As Integer
'        Dim objTarget As acObject
'
'        'Set objTarget = g_Objects.FindObject(pMsg.Value("target"))
'        iStance = pMsg.Value("stance")
'
'        'Spellcasting stance
'        If iStance = &H49 Then
'            myDebug "SpellCasting Anim - iAnimType:" & iAnim1 & " - Atker:" & objEntity.Name ' & " - Target: " & objTarget.Name
'            RaiseEvent OnAttackAnimation(objEntity, objTarget)
'        End If
'
'   ' End If
    
Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnAnimation - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnPlayerKill(pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim objKiller As acObject, objKilled As acObject
    
    Set objKiller = g_Objects.FindPlayer(pMsg.Value("killer"))
    Set objKilled = g_Objects.FindPlayer(pMsg.Value("killed"))
    
    If Valid(objKilled) And Valid(objKiller) Then
        'myDebug "OnPlayerKill: " & objKilled.Name & " was killed by " & objKiller.Name
        objKilled.Dead = True
        Call objKilled.RemoveAllDebuffs
        Call m_GameObjects.FireOnObjectDeath(objKilled)
        RaiseEvent OnPlayerKill(objKiller, objKilled)
    End If
    
Fin:
    Set objKiller = Nothing
    Set objKilled = Nothing
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnPlayerKill - " & Err.Description
    Resume Fin
End Sub


Private Sub NetEvent_OnAddEnchantment(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim iSpellID As Integer
    Dim iLayer As Integer
    Dim lFamily As Long
    Dim dStartTime As Double
    Dim dTimeElapsed As Double
    Dim dDuration As Double
    Dim lSourceGUID As Long
    Dim fAdjust As Single

    With pMsg.Struct("enchantment")
        iSpellID = .Value("spell")
        iLayer = .Value("layer")
        lFamily = .Value("family")
        dTimeElapsed = .Value("elapsedTime")
        dDuration = .Value("duration")
        lSourceGUID = .Value("caster")
        dStartTime = .Value("startTime")
        fAdjust = .Value("value")
    End With
    
    'myDebug "OnAddEnchantment : " & g_Spells.GetSpellName(iSpellId) & " (" & iSpellId & ")"
    'myDebug "OnAddEnchantment: spell=" & g_Spells.GetSpellName(iSpellId) & " (" & iSpellId & "), Family=" & lFamily & ", TimeElapsed=" & dTimeElapsed & "(" & myFormatTime(dTimeElapsed) & "), Duration=" & dDuration & ", startTime=" & dStartTime & ", Source=" & g_Objects.FindObject(lSourceGUID).Name & ", Adjust=" & fAdjust & ", Layer=" & iLayer
    RaiseEvent OnAddEnchantment(iSpellID, lFamily, dStartTime, dTimeElapsed, dDuration, lSourceGUID, iLayer)

Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnAddEnchantment - " & Err.Description
    Resume Fin
End Sub

Private Sub NetEvent_OnRemoveEnchantment(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim iSpellID As Long
    Dim iLayer As Long

    iSpellID = pMsg.Value("spell")
    iLayer = pMsg.Value("layer")

    myDebug "OnRemoveEnchantment: spellId:" & iSpellID & "), Layer=" & iLayer
    
    RaiseEvent OnRemoveEnchantment(iSpellID, iLayer)

Fin:
    Exit Sub
ErrorHandler:
    myError "OnRemoveEnchantment - " & Err.Description
    Resume Fin
End Sub

'NetEvent_OnRemoveMultEnchant
Private Sub NetEvent_OnRemoveMultEnchantment(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim Val As DecalNet.IMessageMember
    Dim mem As DecalNet.IMessageMember

    Dim iSpellID As Long
    Dim iLayer As Long
    Dim iCount As Long
    Dim i As Integer

    iCount = pMsg.Value("count")
    
    myDebug "OnRemoveMultEnchant count: " & iCount
    
    Set Val = pMsg.Struct("enchantments")

    For i = 0 To (iCount - 1)
        iSpellID = Val.Struct(i).Value("spell")
        iLayer = Val.Struct(i).Value("layer")
        RaiseEvent OnRemoveMultEnchantment(iSpellID, iLayer)
    Next i
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnRemoveMultEnchant - " & Err.Description
    Resume Fin
End Sub


Private Sub NetEvent_OnSilentRemoveEnchantment(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim iSpellID As Long
    Dim iLayer As Long

    iSpellID = pMsg.Value("spell")
    iLayer = pMsg.Value("layer")

    myDebug "OnSilentRemoveEnchantment: spellId:" & iSpellID & " Layer=" & iLayer
    
    RaiseEvent OnSilentRemoveEnchantment(iSpellID, iLayer)

Fin:
    Exit Sub
ErrorHandler:
    myError "OnSilentRemoveEnchantment - " & Err.Description
    Resume Fin
End Sub

'NetEvent_OnRemoveMultEnchant
Private Sub NetEvent_OnSilentRemoveMultEnchantment(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim Val As DecalNet.IMessageMember
    Dim mem As DecalNet.IMessageMember

    Dim iSpellID As Long
    Dim iLayer As Long
    Dim iCount As Long
    Dim i As Integer

    iCount = pMsg.Value("count")
    
    myDebug "OnSilentRemoveMultEnchantment count: " & iCount
    
    Set Val = pMsg.Struct("enchantments")

    For i = 0 To (iCount - 1)
        iSpellID = Val.Struct(i).Value("spell")
        iLayer = Val.Struct(i).Value("layer")
        RaiseEvent OnSilentRemoveMultEnchantment(iSpellID, iLayer)
    Next i
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetEvent_OnRemoveMultEnchant - " & Err.Description
    Resume Fin
End Sub



' ****************************************************************************
' *                                Outbound Events                           *
' ****************************************************************************

Private Sub NetAction_CastSpell(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim spellID As Long
    
    spellID = pMsg.Value("spell")
    
    myDebug "NetAction_CastSpell - spellID: " & spellID
    
    RaiseEvent OnSpellCast(spellID)
    
Fin:
    Exit Sub
ErrorHandler:
    myError "NetAction_CastSpell - " & Err.Description
    Resume Fin
End Sub

Private Sub NetAction_CastSpellOn(ByVal pMsg As DecalNet.IMessage2)
On Error GoTo ErrorHandler

    Dim objTarget As acObject
    Dim spellID As Long
    
    spellID = pMsg.Value("spell")
    
    Set objTarget = g_Objects.FindObject(pMsg.Value("target"))
    
    If Valid(objTarget) Then
        myDebug "NetAction_CastSpellOn: spellId " & spellID & " objTarget " & objTarget.Name
        RaiseEvent OnSpellCastOn(spellID, objTarget)
    End If
    
Fin:
    Set objTarget = Nothing
    Exit Sub
ErrorHandler:
    myError "NetAction_CastSpell - " & Err.Description
    Resume Fin
End Sub


