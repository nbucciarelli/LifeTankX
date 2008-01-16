VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type KeyState_t
    Activated As Boolean
    Pushed As Boolean
End Type

Private m_Key(0 To 255) As KeyState_t
Private bMoveHUD As Boolean

'#####################################################################################
'#
'#                         PRIVATE
'#
'#####################################################################################

Private Function IsACHotkey(ByVal KeyNum As Long) As Boolean
    IsACHotkey = (KeyNum >= vbKey0 And KeyNum <= vbKey9)
End Function

'#####################################################################################
'#
'#                          PUBLIC
'#
'#####################################################################################

Public Function HandleKeyboardMessages(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
On Error GoTo ErrorHandler

    Dim bEatKey As Boolean
    Dim i As Integer
    Dim lKey As Long
    Dim x As Long, y As Long
    
    'Default to no message eating
    bEatKey = False
    lKey = wParam
    
    If Not g_bInitComplete Then GoTo Fin
    If g_Hooks.ChatState Then GoTo Fin
    
    Select Case uMsg
        'Key being pushed
        Case WM_KEYDOWN
            If (lKey >= 0) And (lKey <= 255) Then 'valid keyrange
                If (Not m_Key(lKey).Pushed) And Valid(g_Core) Then  'if not already pushed?
                    bEatKey = g_Core.Engine.HandleHotkey(lKey, m_Key(vbKeyControl).Pushed)
                End If
                m_Key(lKey).Pushed = True
            End If
        
        Case WM_KEYUP
            If (lKey >= 0) And (lKey <= 255) Then 'valid keyrange
                m_Key(lKey).Pushed = False
            End If
        Case WM_RBUTTONDOWN
            If Valid(g_Core) Then
                x = GET_X_LPARAM(lParam)
                y = GET_Y_LPARAM(lParam)
                bEatKey = g_Core.Engine.HandleMouseClick(x, y)
                If bEatKey Then
                    bMoveHUD = True
                End If
            End If
        Case WM_LBUTTONDOWN
            If Valid(g_Core) Then
                x = GET_X_LPARAM(lParam)
                y = GET_Y_LPARAM(lParam)
                'MyDebug "HKM: WM_LBUTTONDOWN: x: " & x & " y: " & y

                bEatKey = g_Core.Engine.HandleMouseClick(x, y)
                If bEatKey Then
                    bMoveHUD = True
                End If
            End If
        Case WM_LBUTTONUP
            If bMoveHUD Then
                bMoveHUD = False
                'MyDebug "HKM: WM_LBUTTONUP: x: " & LoWord(lParam) & " y: " & HiWord(lParam)
                If Valid(g_Core) Then
                    x = GET_X_LPARAM(lParam)
                    y = GET_Y_LPARAM(lParam)
                    'MyDebug "HKM: WM_LBUTTONUP: x: " & x & " y: " & y

                    Call g_Core.Engine.HandleMouseMoveHUD(x, y)
                End If
            End If
            
        Case WM_MOUSEMOVE
            'If Valid(g_Core) And bMoveHUD Then
            'End If
            
    End Select
    
Fin:
    HandleKeyboardMessages = bEatKey
    Exit Function
ErrorHandler:
    bEatKey = False
    PrintErrorMessage "HandleKeyboardMessages: " & Err.Description
    Resume Fin
End Function

