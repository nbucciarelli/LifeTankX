VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacroChangeCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Combat Mode change
Option Explicit

Private Enum eSubStates
    ST_START_MODE_CHANGE
    ST_CHECK_MODE_CHANGE
    ST_MODE_CHANGED
End Enum

Private Const TRANSITION_DELAY = 1  'seconds
Private Const SECURE_TIMER = 3      'seconds

Private m_newCombatState As eCombatStates    'the combat mode wanted
Private m_State As eSubStates

Private WithEvents m_tmrSecure As clsTimer
Attribute m_tmrSecure.VB_VarHelpID = -1
Private WithEvents m_tmrTransition As clsTimer
Attribute m_tmrTransition.VB_VarHelpID = -1
Private m_tmrNextThink As clsTimer
Attribute m_tmrNextThink.VB_VarHelpID = -1

Public Event OnCombatModeChanged()

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    Set m_tmrSecure = CreateTimer
    Set m_tmrTransition = CreateTimer
    Set m_tmrNextThink = CreateTimer
    Call Reset
End Sub

Private Sub Reset()
    m_State = ST_START_MODE_CHANGE
    m_newCombatState = COMBATSTATE_NONE
    Call m_tmrSecure.Reset
    Call m_tmrTransition.Reset
    Call m_tmrNextThink.Reset
End Sub

Private Sub Class_Terminate()
    Set m_tmrSecure = Nothing
    Set m_tmrTransition = Nothing
    Set m_tmrNextThink = Nothing
End Sub

'#####################################################################################
'#
'#                                  PROPERTIES
'#
'#####################################################################################

Public Property Get CurrentCombatState() As Integer
    CurrentCombatState = g_Hooks.CombatMode
End Property

'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

Private Sub CombatModeChanged()
    m_tmrTransition.SetNextTime TRANSITION_DELAY   'x seconds delay before giving the hand back to the macro engine
End Sub

'Private Sub m_tmrSecure_OnTimeout()
'    PrintErrorMessage "clsMacroChangeCM.m_tmrSecure_OnTimeout : secure change combat timer expired."
'    Call CombatModeChanged
'End Sub

Private Sub m_tmrTransition_OnTimeout()
    RaiseEvent OnCombatModeChanged
    Call Reset
End Sub

'#####################################################################################
'#
'#                                 PUBLIC
'#
'#####################################################################################

Public Function ChangeCombatState(newCombatState As eCombatStates) As Boolean
    If (CurrentCombatState <> newCombatState) And (newCombatState <> COMBATSTATE_NONE) Then
        m_newCombatState = newCombatState
        m_State = ST_START_MODE_CHANGE
        Call m_tmrSecure.SetNextTime(SECURE_TIMER)
        ChangeCombatState = True
    Else
        ChangeCombatState = False
    End If
End Function

Public Function Restart() As Boolean
On Error GoTo ErrorMessage
    Dim bRet As Boolean
    
    bRet = ChangeCombatState(m_newCombatState)
    
    If Not bRet Then
        Call CombatModeChanged
    End If

Fin:
    Restart = bRet
    Exit Function
ErrorMessage:
    bRet = False
    PrintErrorMessage "clsMacroEquip.Restart - " & Err.Description
    Resume Fin
End Function

Public Sub RunState()
On Error GoTo ErrorMessage

    If m_tmrTransition.Enabled Then Exit Sub
    
    If m_tmrSecure.Expired Then
        MyDebug "clsMacroChangeCM.m_tmrSecure_OnTimeout : secure change combat timer expired."
        Call CombatModeChanged
        Exit Sub
    End If
    
    If m_tmrNextThink.Expired Then
        Select Case m_State
            Case ST_START_MODE_CHANGE
                Call ToggleCombatMode(m_newCombatState)
                Call m_tmrSecure.SetNextTime(SECURE_TIMER)
                Call m_tmrNextThink.SetNextTime(0.75)
                m_State = ST_CHECK_MODE_CHANGE
                
            Case ST_CHECK_MODE_CHANGE
                If CurrentCombatState = m_newCombatState Then   'complete
                    Call m_tmrSecure.Reset
                    Call CombatModeChanged
                Else
                    MyDebug "clsMacroChangeCM.RunState - ST_CHECK_MODE_CHANGE - Combat mode not changed yet, toggling again."
                    Call ToggleCombatMode(m_newCombatState)
                    Call m_tmrNextThink.SetNextTime(0.75)
                End If
        End Select
    End If
    
Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroChangeCM.RunState - " & Err.Description
    Resume Fin
End Sub
