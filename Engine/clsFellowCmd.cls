VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFellowCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Handles the fellowship panel for recruit/dismiss etc

Private Const FELLOW_RECRUIT_XOFF = 250
Private Const FELLOW_RECRUIT_YOFF = 110

Private Const FELLOW_DISBAND_XOFF = 70
Private Const FELLOW_DISBAND_YOFF = 110

Private Const FELLOW_QUIT_XOFF = 160
Private Const FELLOW_QUIT_YOFF = 140

Private Const COMBAT_BAR_HEIGHT = 100
Private Const VITALS_BAR_HEIGHT = 30

Private Const MAX_RECRUIT_RANGE = 15

Private Const MAX_FELLOW = 9

Private Const SECURE_RECRUIT_TIME = 20

Private Enum eSubStates
    FST_NONE
    FST_RECRUIT_START
    FST_RECRUIT_OPEN_PANEL
    FST_RECRUIT_SELECT
    FST_RECRUIT_CLICK_RECRUIT
    FST_RECRUIT_DONE
End Enum

Public Enum eFellowshipAction
    FA_NONE
    FA_RECRUIT
    FA_QUIT
    FA_DISBAND
End Enum

Private WithEvents m_tmrNext As clsTimer
Attribute m_tmrNext.VB_VarHelpID = -1
Private m_tmrSecure As clsTimer

Private m_iAction As eFellowshipAction
Private m_iState As eSubStates

Private m_objTarget As acObject

Public Event OnActionComplete()

'#####################################################################################

Private Sub Class_Initialize()
    Set m_tmrNext = CreateTimer
    Set m_tmrSecure = CreateTimer
End Sub

Private Sub Class_Terminate()
    Set m_tmrNext = Nothing
    Set m_tmrSecure = Nothing
    Set m_objTarget = Nothing
End Sub

Public Sub Reset()
    Call m_tmrNext.Reset
    Call m_tmrSecure.Reset
    Set m_objTarget = Nothing
    m_iAction = FA_NONE
    m_iState = FST_NONE
End Sub

'#####################################################################################

Public Sub ClickRecruit()
    Call Utils.ClickButton(FELLOW_RECRUIT_XOFF, FELLOW_RECRUIT_YOFF)
End Sub

Public Sub ClickDisband()
    Call Utils.ClickButton(FELLOW_DISBAND_XOFF, FELLOW_DISBAND_YOFF)
End Sub

Public Sub ClickQuit()
    Call Utils.ClickButton(FELLOW_QUIT_XOFF, FELLOW_QUIT_YOFF)
End Sub

'#####################################################################################

Public Sub notifyPlayer(ByVal aMsg As String)
    If Valid(m_objTarget) Then
        SendTell m_objTarget.Name, aMsg
    End If
End Sub

Public Sub CheckCommand(ByVal playerName As String, ByVal sMsg As String)
    sMsg = LCase(Trim(sMsg))
    
    MyDebug "cslFellowCmd.CheckCommand: " & playerName & " : (" & sMsg & ")"
    If (sMsg = "list") Then
        Call SendTell(playerName, g_FellowList.getListString)
    ElseIf (sMsg = "remove") Then
        g_FellowList.removeFromLine (playerName)
    End If

End Sub

Public Function RecruitPlayer(objPlayer As acObject) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    'default
    bRet = False
    
    If m_iState = FST_NONE Then
        If ValidRecruit(objPlayer) Then
            Set m_objTarget = objPlayer
            m_iState = FST_RECRUIT_START
            m_iAction = FA_RECRUIT
            Call m_tmrNext.Reset
            Call m_tmrSecure.SetNextTime(SECURE_RECRUIT_TIME)
            bRet = True
        Else
            MyDebug "clsFellowCmd.RecruitPlayer - bad objPlayer recruit"
        End If
    Else
        MyDebug "clsFellowCmd.RecruitPlayer - FellowCmd already busy (m_iState = " & m_iState & ")"
    End If
            
Fin:
    RecruitPlayer = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsFellowCmd.RecruitPlayer - " & Err.Description
    Resume Fin
End Function

Public Sub ActionFinished()
    Call Reset
    RaiseEvent OnActionComplete
End Sub

Private Function ValidRecruit(objPlayer As acObject) As Boolean
On Error GoTo Error_Handler

    Dim bRet As Boolean
    Dim fRange As Single
    
    'default
    bRet = False
    
    If Valid(objPlayer) Then
        'is it truely a player ?
        If g_Objects.IsInPlayersCol(objPlayer) Then
            fRange = objPlayer.GetRange
            
            If fRange > MAX_RECRUIT_RANGE Then
                MyDebug "clsFellowCmd.ValidRecruit - Player out of range"
                Call SendTell(objPlayer.Name, "I can't reach you ! Please stand closer to me.")
            Else
                bRet = True
            End If
        Else
            MyDebug "clsFellowCmd.ValidRecruit - objPlayer is not a valid player object"
        End If
    Else
        MyDebug "clsFellowCmd.ValidRecruit - Invalid objPlayer"
    End If
    
Fin:
    ValidRecruit = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage "clsFellowCmd.ValidRecruit - " & Err.Description
    Resume Fin
End Function

Public Sub OpenFellowPanel()
    g_Core.SendKey vbKeyF12  'make sure we close the fellowship panel (if there was one opened) by opening another panel
    g_Core.SendKey vbKeyF4   'open the fellowshop panel
End Sub

Public Sub RunState()
On Error GoTo ErrorHandler
    
    If m_tmrSecure.Expired Then
        PrintWarning "clsFellowCmd.RunState - Secure timer expired, exiting state."
        Call ActionFinished
        GoTo Fin
    End If
    
    If m_iState <> FST_NONE Then
        If m_tmrNext.Expired Then
            Select Case m_iAction
                Case FA_RECRUIT
                    Select Case m_iState
                        Case FST_RECRUIT_START
                            MyDebug "clsFellowCmd.RunState - RECRUIT - Checking recruit target"
                            If ValidRecruit(m_objTarget) Then
                                Call SendTell(m_objTarget.Name, "Hi " & m_objTarget.Name & " ! I'm going to recruit you, please stand close to me.")
                                Call g_Core.AppActivate  'Make sure AC has focus
                                m_iState = m_iState + 1
                                Call m_tmrNext.SetNextTime(1)
                            Else
                                Call ActionFinished
                            End If
                            
                        Case FST_RECRUIT_OPEN_PANEL
                            MyDebug "clsFellowCmd.RunState - RECRUIT - Opening fellowship panel"
                            Call OpenFellowPanel
                            m_iState = m_iState + 1
                            Call m_tmrNext.SetNextTime(1)
                            
                        Case FST_RECRUIT_SELECT
                            MyDebug "clsFellowCmd.RunState - RECRUIT - Selecting player"
                            Call g_Service.SelectObject(m_objTarget)
                            m_iState = m_iState + 1
                            Call m_tmrNext.SetNextTime(1)
                            
                        Case FST_RECRUIT_CLICK_RECRUIT
                            MyDebug "clsFellowCmd.RunState - RECRUIT - Clicking Recruit button"
                            If ValidRecruit(m_objTarget) Then   'make sure our target didnt move too far away before pushing the button
                                Call ClickRecruit
                                m_iState = m_iState + 1
                                Call m_tmrNext.SetNextTime(4)
                            Else
                                Call ActionFinished
                            End If
                            
                        Case FST_RECRUIT_DONE
                            SendTell m_objTarget.Name, "Ok, you should have been recruited by now. If not, try the command again."
                            Call ActionFinished
                            
                        Case Else
                            'PrintWarning "clsFellowCmd.RunState - unknown recruit step " & m_iState
                            Call ActionFinished
                            
                    End Select
                    
                Case Else
                    'unknown
                    Call ActionFinished
    
            End Select
        End If
    Else
        Call ActionFinished 'no action to perform
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsFellowCmd.RunState - " & Err.Description
    Resume Fin
End Sub

'============================================================================
' PasswordAccepted (PlayerGUID)
'---------------------------------------------------------------------------
' The autofellow recruit password sent by PlayerName is valid
'============================================================================
Public Sub OnPasswordAccepted(playerName As String, PlayerGUID As Long)
    LogEvent playerName & " wants to join the fellowship."
    
    If g_Macro.Ticking Then
        If Not g_Objects.Fellowship.Active Then
            SendTell playerName, "Sorry " & playerName & ", but I'm not part of a fellowship right now."
        ElseIf g_Objects.Fellowship.Exists(PlayerGUID) Then
            SendTell playerName, "Hey, you're already in the fellowship !"
        ElseIf g_FellowList.checkOnList(playerName) Then
            ' Check to see if they are next in line
            If g_FellowList.checkYourTurn(playerName) Then
                Call g_Macro.RecruitPlayer(g_Objects.FindPlayer(PlayerGUID))
            Else
                SendTell playerName, "Sorry, it's not your turn yet"
            End If
        ElseIf (g_FellowList.Count > 0) And g_ui.Macro.chkAutoFellowList.Checked Then
            Call g_FellowList.addToLine(playerName)         ' add to list
        ElseIf g_Objects.Fellowship.NumMembers >= MAX_FELLOW Then
            SendTell playerName, "Sorry " & playerName & ", but the fellowship is currently full (max : " & MAX_FELLOW & ")"
            If g_ui.Macro.chkAutoFellowList.Checked Then
                Call g_FellowList.addToLine(playerName)     ' add to list
            End If
'        ElseIf (g_ui.Fellow.chkRecruitFriendsOnly.Checked = True) And (IsInList(PlayerName, g_ui.Fellow.lstFriends, False) = False) Then
'            SendTell PlayerName, "Sorry " & PlayerName & ", but fellowship recruitment is currently closed."
'            PrintMessage PlayerName & " is not in your friends list - Not recruiting him."
'            Exit Sub
'        ElseIf IsInList(PlayerName, g_ui.Fellow.lstBan, False) = True Then
'            SendTell PlayerName, "Sorry " & PlayerName & ", but you are currently on my Ban list and can't join my fellowship."
'            Exit Sub
        Else
            If Valid(g_Objects.FindPlayer(PlayerGUID)) Then
                Call g_Macro.RecruitPlayer(g_Objects.FindPlayer(PlayerGUID))
            Else
                SendTell playerName, "You are to far away!"
            End If
        End If
        
    Else
        SendTell playerName, "Sorry " & playerName & ", fellowship recruitment is currently unavailable."
    End If
    
End Sub
