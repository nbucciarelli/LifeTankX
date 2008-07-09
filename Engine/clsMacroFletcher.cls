VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacroFletcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
' [[                                                                         [[
' [[                       Fletching State                                   [[
' [[                                                                         [[
' [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

Private Const DEBUG_ME = False
Private Const SECURE_TIMER = 6  'seconds

Private Enum eSubStates
    ST_FLETCH_NONE
    ST_FLETCH_DOFLETCHING
    ST_FLETCH_WAIT
    ST_FLETCH_EQUIPARROW
    ST_FLETCH_COMPLETE
End Enum

Private m_State As eSubStates           'current combat state
Private m_tmrSecure As clsTimer
Private m_tmrNextTime As clsTimer

Private m_sArrow As String
Private m_sArrowHead As String
Private m_sArrowShaft As String

Public Event OnFletchingOver()


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Set m_tmrSecure = CreateTimer
    Set m_tmrNextTime = CreateTimer
    Call Reset
End Sub

Public Sub Reset()
    Call m_tmrSecure.Reset
    Call m_tmrNextTime.Reset
    m_sArrow = ""
    m_sArrowHead = ""
    m_sArrowShaft = ""
    m_State = ST_FLETCH_NONE
End Sub

Private Sub Class_Terminate()
    Set m_tmrSecure = Nothing
    Set m_tmrNextTime = Nothing
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   PRIVATE
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Function ValidFletchItem(ByVal sItemName As String) As Boolean
    ValidFletchItem = ((Trim(sItemName) <> "") And (sItemName <> TAG_NOT_SET))
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                External Events
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Sub OnFletchingSuccessful()
    locDebug "Fletching successful, equipping arrow"
    LogEvent "Fletching successful."
    m_State = ST_FLETCH_EQUIPARROW
    Call m_tmrNextTime.SetNextTime(0.6)
End Sub

Public Sub OnFletchingFailed()
    locDebug "Fletching failed, trying again"
    If Not StartFletching("OnFletchingFailed") Then
        Call StopFletching("OnFletchingFailed")
    End If
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   PUBLIC
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function StartFletching(Optional ByVal sSource As String = "") As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    Dim sArrowName As String
    Dim sArrowHead As String
    Dim sArrowShaft As String
    
    locDebug "StartFletching - Src: " & sSource

    sArrowName = g_Data.GetArrowType
    sArrowHead = g_Data.GetArrowHead
    sArrowShaft = g_Data.ArrowShaft

    If ValidFletchItem(sArrowName) _
    And ValidFletchItem(sArrowHead) _
    And ValidFletchItem(sArrowShaft) Then
        Call Reset
        m_sArrow = sArrowName
        m_sArrowHead = sArrowHead
        m_sArrowShaft = sArrowShaft
        m_State = ST_FLETCH_DOFLETCHING
        Call m_tmrSecure.SetNextTime(SECURE_TIMER)
        Call m_tmrNextTime.SetNextTime(1)
        bRet = True
    Else
        bRet = False
    End If
    
Fin:
    StartFletching = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsMacroFletcher.StartFletching - " & Err.Description
    Resume Fin
End Function

Public Function RestartFletching() As Boolean
    RestartFletching = StartFletching(RestartFletching)
End Function

Public Function StopFletching(Optional ByVal sSource As String = "")
    locDebug "StopFletching - Src: " & sSource
    Call Reset
    RaiseEvent OnFletchingOver
End Function

Public Sub EquipAvailableArrows()
On Error GoTo ErrorHandler
    Dim objArrows As acObject
    
    Set objArrows = g_Objects.Items.InvFindByName(m_sArrow)
    If (objArrows.Guid <> -1) And (Not objArrows.Equiped) Then
        Call g_Macro.RequestEquipItem(objArrows) 'g_Service.UseItemOnSelf(objArrows)
    End If
    
Fin:
    Set objArrows = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacroFletcher.EquipAvailableArrows - " & Err.Description
    Resume Fin
End Sub

'******************************************
' RunState
'
' Handles the fletching states
'******************************************
Public Sub RunState()
On Error GoTo ErrorHandler

    Dim iNumHeads As Integer
    Dim iNumShafts As Integer
    Dim objArrows As acObject
    
    If m_tmrSecure.Expired Then
        PrintWarning "clsMacroFletcher Secure Timer - Secure timer expired, exiting state"
        Call StopFletching("clsMacroFletcher Secure Timer")
        Exit Sub
    End If
    
    If Not m_tmrNextTime.Expired Then Exit Sub
    
    Select Case m_State

        Case ST_FLETCH_DOFLETCHING
            If g_Hooks.CombatMode <> COMBATSTATE_PEACE Then
                Call g_Macro.RequestCombatStateChange(COMBATSTATE_PEACE)
                Exit Sub
            End If
            
            iNumHeads = g_Objects.Items.InvCntByName(m_sArrowHead, True)
            iNumShafts = g_Objects.Items.InvCntByName(m_sArrowShaft, True)
            
            If iNumHeads > 0 And iNumShafts > 0 Then
                Call g_Service.UseItemOn(g_Objects.Items.InvFindByName(m_sArrowHead), g_Objects.Items.InvFindByName(m_sArrowShaft))
                m_State = ST_FLETCH_WAIT    'wait for OnFletchingSuccessful/Failed event
            Else
                LogEvent "Out of fletching components (" & iNumHeads & " " & m_sArrowHead & ", " & iNumShafts & " " & m_sArrowShaft & " left)"
                Call StopFletching("RunState - Out of fletching components")
            End If
        
        Case ST_FLETCH_EQUIPARROW
            locDebug "ST_FLETCH_EQUIPARROW - Trying to equip the fletched arrows..."
            Call EquipAvailableArrows
            m_State = ST_FLETCH_COMPLETE
            Call m_tmrNextTime.SetNextTime(1)
            
        Case ST_FLETCH_COMPLETE
            Call StopFletching("ST_FLETCH_COMPLETE")
            
    End Select
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsMacroFletcher.RunState - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Local Utility Functions
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function GetStateString() As String
Dim sRes As String

    Select Case m_State
            
        Case ST_FLETCH_DOFLETCHING
            sRes = "Doing Fletching"
            
        Case ST_FLETCH_EQUIPARROW
            sRes = "Equipping Arrow"
            
        Case ST_FLETCH_COMPLETE
            sRes = "Fletch Complete"
            
        Case ST_FLETCH_WAIT
            sRes = "Waiting"
      
        Case Else
            sRes = "Unknown"
    End Select
    
    GetStateString = sRes
End Function

'Local Debug
Private Sub locDebug(DebugMsg As String, Optional bSilent As Boolean = True)
    If DEBUG_ME Or g_Data.mDebugMode Then
        Call MyDebug("[clsMacroFletcher : " & GetStateString & "] " & DebugMsg, bSilent)
    End If
End Sub

