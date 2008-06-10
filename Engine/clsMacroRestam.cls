VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacroRestam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Melee Restam  State                               [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False

Private Enum eSubStates
    STATE_RESTAM_READY
    STATE_RESTAMING
    STATE_DONE
End Enum

Public Enum eRestamMethod
    RESTAM_SPELL
    RESTAM_ITEM
    NO_RESTAM
End Enum

Private Const TRANSITION_DELAY = 0.5    'seconds
Private Const SECURE_TIME = 2.5         'seconds

Private WithEvents m_tmrTransition As clsTimer
Attribute m_tmrTransition.VB_VarHelpID = -1
Private WithEvents m_tmrSecure As clsTimer
Attribute m_tmrSecure.VB_VarHelpID = -1

Private m_RestamMethod As eRestamMethod
Private m_State As eSubStates
Private m_tmrNextTime As clsTimer

Public Event OnRestamComplete()


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Set m_tmrTransition = CreateTimer
    Set m_tmrSecure = CreateTimer
    Set m_tmrNextTime = CreateTimer
    Call Reset
End Sub

Public Sub Reset()
    m_RestamMethod = RESTAM_SPELL
    m_State = STATE_RESTAM_READY
    Call m_tmrSecure.Reset
    Call m_tmrTransition.Reset
    Call m_tmrNextTime.Reset
End Sub

Private Sub Class_Terminate()
    Set m_tmrTransition = Nothing
    Set m_tmrSecure = Nothing
    Set m_tmrNextTime = Nothing
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get RestamMethod() As eRestamMethod
    RestamMethod = m_RestamMethod
End Property


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Private Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub m_tmrTransition_OnTimeout()
    If m_State = STATE_DONE Then
        locDebug "clsMacroRestam.m_tmrTransition_OnTimeout: STATE_DONE, so calling OnRestamComplete"
        RaiseEvent OnRestamComplete
    End If
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Public Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Sub ChooseBestRestamMethod()
On Error GoTo ErrorMessage
    
    If m_RestamMethod = NO_RESTAM Then
        Exit Sub
    End If
    
    ' First see if we can use Stam Items
    If Vitals.CanUseStamItem Then
        m_RestamMethod = RESTAM_ITEM
    Else
        ' Or can we use spells?
        If g_ui.Macro.chkUseStamSpell.Checked Then
            m_RestamMethod = RESTAM_SPELL
        Else
            PrintWarning "Trying to restam, but no stamina items left! Doing nothing... Either enable spells or restock!"
            m_RestamMethod = NO_RESTAM
            Exit Sub
        End If
    End If
    
Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroRestam.RunState - " & Err.Description
    Resume Fin
End Sub

Public Function StartRestam() As Boolean
On Error GoTo ErrorMessage
    
    Dim bRet As Boolean
    bRet = False

    locDebug "StartRestam..."
    
    m_State = STATE_RESTAM_READY
    Call m_tmrSecure.SetNextTime(SECURE_TIME)
    Call g_Spells.ResetCastingFlag("StartRestam")
    
    bRet = True

Fin:
    StartRestam = bRet
    Exit Function
ErrorMessage:
    bRet = False
    PrintErrorMessage "clsMacroRestam.StartRestam - " & Err.Description
    Resume Fin
End Function

Public Sub StopRestam(Optional ByVal sSource As String = "")
On Error GoTo ErrorMessage

    locDebug "clsMacroRestam.StopRestam(" & sSource & ")"
    Call m_tmrSecure.Reset
    Call m_tmrTransition.SetNextTime(TRANSITION_DELAY)
    m_State = STATE_DONE

Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroRestam.StopRestam - " & Err.Description
    Resume Fin
End Sub

'RunState
Public Sub RunState()
On Error GoTo ErrorMessage
    Dim objPotion As acObject
    Dim bNeedStam As Boolean
    
    'Must wait TRANSITION_TIME
    If m_tmrTransition.Enabled Then Exit Sub
    
    'check if we still need to restam
    If g_Filters.Stam > GetPercent(g_Filters.MaxStam, RESTAM_MAX_STAM_PERCENT) Then
        Call StopRestam("All Stammed up!")
        Exit Sub
    End If
    
    If m_tmrSecure.Expired Then
        locDebug "clsMacroRestam.m_tmrSecure_OnTimeout - Secure Restam Timer expired : exiting state."
        locDebug "Busy Timeout - Reseting Restam flags"
        Call StopRestam("Secure Restam Timer Expired")
        Exit Sub
    End If
 
    Select Case m_State
    
        Case STATE_RESTAM_READY
                'Select a restam method
                Call ChooseBestRestamMethod
                
                Select Case m_RestamMethod
                    Case RESTAM_ITEM
                        If (g_bFindingItem) Then GoTo Fin
                        Set objPotion = Vitals.findStamItem
                        If Valid(objPotion) And m_tmrNextTime.Expired Then
                            MyDebug "clsMacroRestam.RunState: Using " & objPotion.Name
                            Call g_Service.UseItem(objPotion) ' UseItemOnSelf
                            Call m_tmrNextTime.SetNextTime(0.5)
                            GoTo Fin
                        Else
                            MyDebug "clsMacroRestam.RunState: Stamina timer not expired!"
                            GoTo Fin
                        End If
                        
                    Case RESTAM_SPELL   'spell
                        If g_Macro.ValidState(TYPE_CASTER) Then
                            If Not Vitals.CastManaStamRegen Then    'Juggle mana
                                locDebug "clsMacroRestam.RunState: Casting Revitalize"
                                Call g_Spells.Cast_Revitalize
                                'move to "restaming state"
                                m_State = STATE_RESTAMING
                            End If
                        Else
                            'Exit until we're in valid casting state
                            Call m_tmrSecure.SetNextTime(SECURE_TIME)
                            GoTo Fin
                        End If
                        
                    Case Else 'No restam.
                        Call StopRestam("STATE_DONE")
                        
                End Select
             
        Case STATE_RESTAMING
            'do nothing
            If m_RestamMethod = RESTAM_SPELL And Not (g_Spells.Casting) Then
                MyDebug "clsMacroRestam.RunState:STATE_RESTAMING: spamming the spell"
                Call g_Spells.Cast_Revitalize
            End If
            
        Case STATE_DONE
            ' We are all done
            Call StopRestam("STATE_DONE")
            
    End Select
    
Fin:
    Set objPotion = Nothing
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroRestam.RunState - " & Err.Description
    Resume Fin
End Sub

'Previous restam action complete
Public Sub OnRestamReady(Optional ByVal bSuccess As Boolean = True)
    If (m_State = STATE_RESTAMING) Or (Not bSuccess) Then
        If Not m_tmrTransition.Enabled Then
            If bSuccess Then
                MyDebug "OnRestamReady - Revit Spell/Potion used"
            End If
            
            Call m_tmrSecure.SetNextTime(SECURE_TIME)
            m_State = STATE_RESTAM_READY
        End If
    End If
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function GetCurrentStateString() As String
Dim sRes As String

    sRes = ""
    
    Select Case m_RestamMethod

        Case RESTAM_SPELL
            sRes = sRes & "[Spell]"
            
        Case RESTAM_ITEM
            sRes = sRes & "[Item]"
            
        Case Else
            sRes = sRes & "[Unknown Restam Method " & m_RestamMethod & "]"
            
    End Select
    
    GetCurrentStateString = sRes
End Function


'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsMacroRestam] " & DebugMsg, bSilent)
    End If
End Sub

