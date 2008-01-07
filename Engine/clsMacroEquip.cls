VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMacroEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eSubStates
    ST_EQUIP_START
    ST_EQUIP_RUNNING
End Enum

Private Const TRANSITION_DELAY = 0.5     'seconds
Private Const SECURE_TIME = 5            'seconds
Private Const RETRY_DELAY = 2            'seconds

Private m_State As eSubStates
Private m_ItemToEquip As acObject

Private WithEvents m_tmrTransition As clsTimer
Attribute m_tmrTransition.VB_VarHelpID = -1
Private WithEvents m_tmrSecure As clsTimer
Attribute m_tmrSecure.VB_VarHelpID = -1
Private WithEvents m_tmrRetry As clsTimer
Attribute m_tmrRetry.VB_VarHelpID = -1

Public Event OnItemEquipped()

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    Set m_tmrTransition = CreateTimer
    Set m_tmrSecure = CreateTimer
    Set m_tmrRetry = CreateTimer
    Call Reset
End Sub

Private Sub Reset()
    m_State = ST_EQUIP_START
    Set m_ItemToEquip = Nothing
    Call m_tmrTransition.Reset
    Call m_tmrSecure.Reset
    Call m_tmrRetry.Reset
End Sub

Private Sub Class_Terminate()
    Set m_ItemToEquip = Nothing
    Set m_tmrTransition = Nothing
    Set m_tmrSecure = Nothing
    Set m_tmrRetry = Nothing
End Sub



'#####################################################################################
'#
'#                                  PROPERTIES
'#
'#####################################################################################


'#####################################################################################
'#
'#                                  PRIVATE
'#
'#####################################################################################

'Private Sub m_tmrSecure_OnTimeout()
'    PrintWarning "clsMacroEquip.SecureTimer : expired - couldn't equip the item. Stopping."
'    Call StopEquip
'End Sub

Private Sub m_tmrTransition_OnTimeout()
    MyDebug "Equip Transition Over - Firing Event"
    RaiseEvent OnItemEquipped
    Call Reset
End Sub

Private Sub StopEquip()
    Call m_tmrTransition.SetNextTime(TRANSITION_DELAY)
    MyDebug "StopEquip - Setting next time to " & m_tmrTransition.Value
End Sub

'#####################################################################################
'#
'#                                 PUBLIC
'#
'#####################################################################################

Public Function StartEquip(objItem As acObject) As Boolean
On Error GoTo ErrorMessage
    
    Dim bRet As Boolean
    bRet = False
    Set m_ItemToEquip = Nothing

    'Make sure we have something to equip
    If Not Valid(objItem) Then
        PrintWarning "clsMacroEquip.StartEquip : objItem = NULL. Exiting."
        GoTo Fin
    End If
    
    'Make sure it's in our inventory
    If Not g_Objects.Items.InInventory(objItem.Guid) Then
        PrintWarning "clsMacroEquip.StartEquip : " & objItem.Name & " not in inventory. Exiting."
        GoTo Fin
    ElseIf objItem.Equiped Then
        PrintWarning "clsMacroEquip.StartEquip : " & objItem.Name & " is already equipped. Exiting."
        GoTo Fin
    End If
    
    MyDebug "(clsMacroEquip) Equipping : " & objItem.Name & "   :Equiped:" & objItem.Equiped
    Set m_ItemToEquip = objItem
    m_State = ST_EQUIP_START
    Call m_tmrSecure.SetNextTime(SECURE_TIME)
    bRet = True

Fin:
    StartEquip = bRet
    Exit Function
ErrorMessage:
    bRet = False
    PrintErrorMessage "clsMacroEquip.StartEquip - " & Err.Description
    Resume Fin
End Function

Public Function Restart() As Boolean
On Error GoTo ErrorMessage
    Dim bRet As Boolean
    
    bRet = StartEquip(m_ItemToEquip)
    
    If Not bRet Then
        Call StopEquip
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
        PrintWarning "clsMacroEquip.SecureTimer : expired - couldn't equip the item. Stopping."
        Call StopEquip
        Exit Sub
    End If
    
    Select Case m_State
    
        Case ST_EQUIP_START
            Call g_Service.EquipItem(m_ItemToEquip)
            Call m_tmrSecure.SetNextTime(SECURE_TIME)
            Call m_tmrRetry.SetNextTime(RETRY_DELAY)
            m_State = ST_EQUIP_RUNNING
            
        Case ST_EQUIP_RUNNING
            If m_ItemToEquip.Equiped Then
                MyDebug m_ItemToEquip.Name & " equiped !"
                Call StopEquip
            Else
                MyDebug m_ItemToEquip.Name & " still not equiped..."
                If m_tmrRetry.Expired Then
                    Call g_Service.EquipItem(m_ItemToEquip)
                    Call m_tmrRetry.SetNextTime(RETRY_DELAY)
                    MyDebug "clsMacroEquip: tmrRetry timed out, calling EquipItem again"
                End If
            End If
            
    End Select
    
Fin:
    Exit Sub
ErrorMessage:
    PrintErrorMessage "clsMacroEquip.RunState - " & Err.Description
    Call StopEquip
    Resume Fin
End Sub


