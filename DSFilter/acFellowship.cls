VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acFellowship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Fellowship class
Option Explicit

Private m_colFellow As colObjects   'the collection of players in the fellowship
                                    'IMPORTANT NOTE : this only holds the player NAMES & GUIDS,
                                    'and not a pointer to the actual object
                                    'The reason behind this is that we may recieve the name & guid
                                    'of players we aren't aware of yet
                                    'NOTE : this collection also includes a record with our GUID/Name
                                    
Private m_objLeader As acObject    'the fellowship leader - points to a member of m_colFellow (name/id info)
                                    
Private m_sFellowshipName As String 'name of the fellowship
Private m_bShareLoot As Boolean     'does the fellowship supports loot sharing?
Private m_bInFellowship As Boolean  'flag to tell if we're already part of a fellowship or not

Public Event OnCreate()     'we're creating a fellowship
Public Event OnJoin()       'we're joining a fellowship
Public Event OnRecruit(ByVal objFellow As acObject)  'someone else got recruted in fellowship
Public Event OnQuit(ByVal objFellow As acObject)    'someone left the fellow (can be us)
Public Event OnKick(ByVal objFellow As acObject)  'someone got dismissed from fellow (can be us)
Public Event OnDisband()  'the fellowship got disbanded
Public Event OnRecvInvite()     'someone's inviting us
Public Event OnUpdateFellowStats(ByVal objPlayer As acObject)

'============================================================
'                Constructor/Destructor
'============================================================

Private Sub Class_Initialize()
    Call ResetFellowship
End Sub

Private Sub Class_Terminate()
    Set m_colFellow = Nothing
    Set m_objLeader = Nothing
End Sub

Public Sub ResetFellowship()
    Set m_colFellow = New colObjects
        m_colFellow.Description = "Fellowship"
    Set m_objLeader = Nothing
    m_sFellowshipName = "UnknownFellow"
    m_bShareLoot = False
    m_bInFellowship = False
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colFellow.NewEnum
End Function

'============================================================
'                       Properties
'============================================================

Public Property Get Name() As String
    Name = m_sFellowshipName
End Property

Public Property Let Name(ByVal sVal As String)
    m_sFellowshipName = sVal
End Property

Public Property Get NumMembers() As Long
    NumMembers = m_colFellow.Count
End Property

Public Property Get ShareLoot() As Boolean
    ShareLoot = m_bShareLoot
End Property

Public Property Get Leader() As acObject
    Set Leader = m_objLeader
End Property

'is fellowship active?
Public Property Get Active() As Boolean
    Active = m_bInFellowship
End Property

'============================================================
'                       Private
'============================================================

Private Function AddFellow(ByVal sPlayerName As String, ByVal lPlayerGUID As Long, Optional bShareLoot As Boolean = False) As acObject
On Error GoTo ErrorHandler
    Dim objFellow As acObject
    
    'add player name/guid info to our fellow collection
    Set objFellow = m_colFellow.Add(lPlayerGUID, sPlayerName)

    'share loot?
    objFellow.ShareLoot = bShareLoot
    
Fin:
    Set AddFellow = objFellow
    Exit Function
ErrorHandler:
    Set objFellow = Nothing
    'myError "acFellowship.AddFellow(" & sPlayerName & ", " & lPlayerGUID & ")"
    Resume Fin
End Function

'============================================================
'                       Public
'============================================================

Public Function Exists(ByVal lGUID As Long) As Boolean
    Exists = m_colFellow.Exists(lGUID)
End Function

Public Function FindFellow(ByVal lGUID As Long) As acObject
On Error GoTo ErrorHandler

    If m_colFellow.Exists(lGUID) Then
        Set FindFellow = m_colFellow(lGUID)
    Else
        Set FindFellow = Nothing
    End If
    
Fin:
    Exit Function
ErrorHandler:
    'myError "acFellowship.FindFellow(" & lGUID & ")"
    Set FindFellow = Nothing
    Resume Fin
End Function

Public Function FindFellowByName(ByVal sName As String) As acObject
On Error GoTo ErrorHandler

    Dim objFellow As acObject
    Dim objSelected As acObject
    
    Set objSelected = Nothing
    For Each objFellow In m_colFellow
        If SameText(objFellow.Name, sName) Then
            Set objSelected = objFellow
            GoTo Fin
        End If
    Next objFellow
    
Fin:
    Set FindFellowByName = objSelected
    Set objFellow = Nothing
    Set objSelected = Nothing
    Exit Function
ErrorHandler:
    'myError "acFellowship.FindFellowByName(" & sName & ")"
    Set objSelected = Nothing
    Resume Fin
End Function

Public Function PlayerShareLoot(ByVal sName As String) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim objFellow As acObject
    
    Set objFellow = FindFellowByName(sName)
    If Valid(objFellow) Then
        bRet = objFellow.ShareLoot
    Else
        bRet = False
    End If
    
Fin:
    PlayerShareLoot = bRet
    Exit Function
ErrorHandler:
    'myError "acFellowship.PlayerShareLoot(" & sName & ")"
    bRet = False
    Resume Fin
End Function

'============================================================================
' HandleFellowshipEvents
'---------------------------------------------------------------------------
' Called when receiving fellowship messages from the server
'
' Handles what to do when receiving a fellowship create/recruit/dismiss/disband
' request
'
'============================================================================
Public Sub HandleFellowshipEvents(ByVal pMsg As DecalNet.IMessage2, ByVal lFellowEvent As Long)
On Error GoTo Error_Handler

    Dim lPlayerGUID As Long, lFellowLeaderGUID As Long
    Dim sFellowshipName As String, sPlayerName As String
    Dim iFellowCount As Integer, i As Integer
    Dim bShareLoot As Boolean
    Dim Val As DecalNet.IMessageMember
    Dim mem As DecalNet.IMessageMember
    
    Dim objFellow As acObject      'holds a fellow member GUID/Name info
    Dim objPlayer As acObject      'points to a valid/known player object
    
    Select Case lFellowEvent
    
        '[[[ Fellowship Creation ]]]
        Case EV_FELLOWSHIP_CREATE
            myDebug "HandleFellowshipEvents: Fellowship_Create Message"
            
            'If we were not in a fellowship already, it means we're joining/creating one
            If Not m_bInFellowship Then
                
                'reset our fellows collection and fellow count
                Call ResetFellowship
                
                'get data from message
                iFellowCount = CInt(pMsg.Value("fellowCount"))
                
                myDebug "FellowCount = " & iFellowCount
                
                m_sFellowshipName = pMsg.Value("name")
                lFellowLeaderGUID = pMsg.Value("leader")
                
                myDebug "Fellowship: name: " & m_sFellowshipName & "   leader: " & lFellowLeaderGUID
                
                'We're now part of the fellowship
                m_bInFellowship = True
                
                Set Val = pMsg.Struct("fellows")
                
                'Now cycle through all the fellowship members and save their guid in
                'the fellow collection (this includes the fellow Leader as well)
                For i = 0 To (iFellowCount - 1)
                     'Set mem = pMsg.Struct("fellows").Struct(i).Struct("fellow")
                    Set mem = Val.Struct(i).Struct("fellow")
                    
                    'current fellow member guid
                    lPlayerGUID = mem.Value("fellow")
                    sPlayerName = CStr(mem.Value("name"))
                    bShareLoot = (mem.Value("shareLoot") = 1)
                        
                    myDebug "..." & sPlayerName
                        
                    'add player name/guid info to our fellow collection
                    Set objFellow = AddFellow(sPlayerName, lPlayerGUID, bShareLoot)
      
                    'set the fellowship leader object based on the guid we just grabbed
                    If objFellow.GUID = lFellowLeaderGUID Then
                        Set m_objLeader = objFellow
                    End If
                        
                    'Only try to update stats of players we are aware of
                    '(that is, player objects existing
                    'in our g_object.Players collection)
                    'Set objPlayer = g_Objects.FindPlayer(lPlayerGUID)
                    If Valid(objFellow) Then
                        objFellow.MaxHealth = mem.Value("maxHealth")
                        objFellow.MaxStamina = mem.Value("maxStam")
                        objFellow.MaxMana = mem.Value("maxMana")
                            
                        objFellow.Health = mem.Value("curHealth")
                        objFellow.Stamina = mem.Value("curStam")
                        objFellow.Mana = mem.Value("curMana")
                            
                        RaiseEvent OnUpdateFellowStats(objFellow)
                    End If
                Next i
                    
                
                'debug
                myDebug "Fellow leader : " & m_objLeader.Name
                
                'If we're the fellowship creator/leader...
                If g_Objects.IsSelf(m_objLeader) Then
                    myDebug "Creating fellowship..."
                    RaiseEvent OnCreate
                Else
                    myDebug "Joining fellowship..."
                    RaiseEvent OnJoin
                End If
                
            Else    'we were already in a fellowship
                    ' do nothing
            End If
           
        '[[[ Quit Fellowship / Dismiss Member ]]]
        Case EV_FELLOWSHIP_QUIT, EV_FELLOWSHIP_DISMISS
            myDebug "OnFellowshipEvents : fellowship quit/dismiss"
            
            'guid of player who's quitting or getting dismissed
            lPlayerGUID = pMsg.Value("fellow")
            Set objFellow = FindFellow(lPlayerGUID)
            
            If Valid(objFellow) Then
                If lFellowEvent = EV_FELLOWSHIP_QUIT Then
                    RaiseEvent OnQuit(objFellow)
                Else
                    RaiseEvent OnKick(objFellow)
                End If
                
                If g_Objects.IsSelf(objFellow) Then
                    myDebug "FellowshipQuit/Dismiss: Macro left fellow"
                    Call ResetFellowship    'clear fellowship data
                Else
                    Call m_colFellow.Remove(lPlayerGUID)
                End If
            Else
                'myError "FellowshipQuit/Dismiss : unknown player #" & lPlayerGUID & " left the fellow."
            End If
        
        '[[[ Fellowship gets dibanded ]]]
        Case EV_FELLOWSHIP_DISBANDS
            myDebug "HandleFellowshipEvents : fellowship disbands"
            
            'fire the event
            RaiseEvent OnDisband
    
            'clear fellowship info
            Call ResetFellowship
            
        '[[[ Fellowship Recruit ]]]
        'Note : since March patch, the fellowship recruit message is also used to
        'update the fellowship members' stats
        Case EV_FELLOWSHIP_RECRUIT
            'myError "OnFellowshipEvents : fellowship recruit/update message"
            
            Set Val = pMsg.Struct("fellow")

            lPlayerGUID = Val.Value("fellow")
            sPlayerName = Val.Value("name")
            bShareLoot = (Val.Value("shareLoot") = 1)
            
            Set objFellow = FindFellow(lPlayerGUID)
                
            If Not Valid(objFellow) Then    'player not in fellowship yet
                Set objFellow = AddFellow(sPlayerName, lPlayerGUID, bShareLoot)
                If Valid(objFellow) Then
                    myDebug "OnFellowshipEvents : firing OnRecruit"
                    RaiseEvent OnRecruit(objFellow)
                Else
                    'myError "HandleFellowshipEvents - EV_FELLOWSHIP_RECRUIT : invalid objFellow"
                    GoTo Fin
                End If
            End If
            
            'Fellowship Members Stats Update
            'Only try to update stats of players we are aware of
            '(that is, player objects existing in our g_object.Players collection)
            'Set objPlayer = g_Objects.FindPlayer(lPlayerGUID)
            If Valid(objFellow) Then
                objFellow.MaxHealth = Val.Value("maxHealth")
                objFellow.MaxStamina = Val.Value("maxStam")
                objFellow.MaxMana = Val.Value("maxMana")
                
                objFellow.Health = Val.Value("curHealth")
                objFellow.Stamina = Val.Value("curStam")
                objFellow.Mana = Val.Value("curMana")
                
                RaiseEvent OnUpdateFellowStats(objFellow)
            End If
               
        Case RETIRED_EV_FELLOWSHIP_INVITATION
            RaiseEvent OnRecvInvite
            GoTo Fin
        
        Case Else
            myDebug "HandleFellowshipEvents: unkown FellowEvent " & lFellowEvent
            GoTo Fin
            
    End Select
    
Fin:
    Set objFellow = Nothing
    Set objPlayer = Nothing
    Exit Sub
Error_Handler:
    'myError "clsACObjects.HandleFellowshipEvents (" & lFellowEvent & ") - " & Err.Description & " - line: " & Erl
    Resume Fin
End Sub

