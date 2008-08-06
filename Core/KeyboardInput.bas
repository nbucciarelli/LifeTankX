Attribute VB_Name = "Inputs"
Option Explicit


Public Function mySendMessage(ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
    If g_PluginSite.hwnd <> 0 Then
        'Call SendMessage(g_PluginSite.hwnd, WM_SETFOCUS, 0, ByVal 0&)
        mySendMessage = SendMessage(g_PluginSite.hwnd, ByVal wMsg, ByVal wParam, ByVal lParam)        'modified : Byval lParam
    End If
End Function

Public Function myPostMessage(ByVal wMsg As Long, ByVal wParam As Long, lParam As Long)
    If g_PluginSite.hwnd <> 0 Then
        myPostMessage = PostMessage(g_PluginSite.hwnd, wMsg, wParam, lParam)
    End If
End Function

Public Function myAppActivate()
    Call mySendMessage(WM_ACTIVATE, WA_ACTIVATE, &H0)
    ' WM_ACTIVATEAPP = &H1C
    'Call mySendMessage(&H1C, &H1, &H0)
    'WM_NCACTIVATE = &H86
    'Call mySendMessage(&H86, &H1, &H0)
End Function

Public Function IsExtendedKey(ByVal lKey As Long) As Boolean
    Select Case lKey
        Case vbKeyMenu, _
          vbKeyInsert, vbKeyDelete, vbKeyHome, vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
          vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, _
          vbKeyDivide, vbKeyExecute, vbKeySnapshot ' vbkeycontrol ?
            IsExtendedKey = True
        Case Else
            IsExtendedKey = False
    End Select
End Function


Public Sub myAppSendKey(ByVal Key As Long)
Dim lParam As Long

    Call myAppActivate
    
    lParam = 1  ' Repeat count.
    lParam = lParam Or MapVirtualKey(Key, 0) * &H10000   ' Scan code
    
    If IsExtendedKey(Key) Then
        lParam = lParam Or &H1000000
    End If
    
    Call mySendMessage(WM_KEYDOWN, Key, lParam)
    lParam = lParam Or &HC0000000  ' Previous key state & transition state
    Call mySendMessage(WM_KEYUP, Key, lParam)
End Sub

Public Sub myAppSendKeyHold(ByVal Key As Long)
Dim lParam As Long

    Call myAppActivate
    
    lParam = 1  ' Repeat count.
    lParam = lParam Or MapVirtualKey(Key, 0) * &H10000  ' Scan code
    If IsExtendedKey(Key) Then lParam = lParam Or &H1000000
    Call mySendMessage(WM_KEYDOWN, Key, lParam)

End Sub

Public Sub myAppSendKeyRelease(ByVal Key As Long)
Dim lParam As Long

    Call myAppActivate
    
    lParam = 1  ' Repeat count.
    lParam = lParam Or MapVirtualKey(Key, 0) * &H10000 Or &HC0000000
    If IsExtendedKey(Key) Then lParam = lParam Or &H1000000
    Call mySendMessage(WM_KEYUP, Key, lParam)

End Sub

'Send char (caught by AC console)
Public Sub myAppSendChar(Key As Long)
Dim lParam As Long

    Call myAppActivate
    
    lParam = 1  ' Repeat count.
    lParam = lParam Or MapVirtualKey(Key, 0) * &H10000  ' Scan code
    Call mySendMessage(WM_KEYDOWN, Key, lParam)
    Call mySendMessage(WM_CHAR, Key, lParam)
    lParam = lParam Or &HC0000000  ' Previous key state & transition state
    Call mySendMessage(WM_KEYUP, Key, lParam)
    
End Sub

Public Sub myAppSendPaste()

    myAppSendKeyHold vbKeyControl
    myAppSendKeyHold vbKeyV
    
    myAppSendKeyRelease vbKeyV
    myAppSendKeyRelease vbKeyControl

End Sub

Public Sub mySendTextToConsole(sText As String, Optional bForceSend As Boolean = False, Optional SendMessage As Boolean = True)
    
    If g_Hooks.ChatState Then
        If bForceSend Then
            myAppSendKey vbKeyReturn
        Else
            PrintMessage "Tryed to send message : '" & sText & "', but you were already chatting."
            Exit Sub
        End If
    End If
    
    ' Call g_Hooks.InvokeChatParser(sText)

    
    Call Clipboard.Clear
    Call Clipboard.SetText(sText, 1)
    
    myAppSendKey vbKeyReturn
    myAppSendPaste
    If SendMessage Then myAppSendKey vbKeyReturn
    
End Sub

'This function cant be used to send input to the the AC console. Use mySendKeyEx instead
Public Sub mySendKey(ByVal Key As Long)

    If g_PluginSite.Hooks.ChatState = True Then
        'MyDebug ">>> TAB key mode"
        'Tab, send our key, and tab again
        myAppSendKey vbKeyTab
        myAppSendKey Key
        myAppSendKey vbKeyTab
    Else
        myAppSendKey Key
    End If

End Sub

Public Sub mySendKeyHold(Key As Long)

    If g_Hooks.ChatState = True Then
        'MyDebug ">>> TAB key mode"
        'Tab, send our key, and tab again
        myAppSendKey vbKeyTab
        myAppSendKeyHold Key
        myAppSendKey vbKeyTab
    Else
        myAppSendKeyHold Key
    End If

End Sub

Public Sub mySendKeyRelease(Key As Long)

    If g_Hooks.ChatState = True Then
        'MyDebug ">>> TAB key mode"
        'Tab, send our key, and tab again
        myAppSendKey vbKeyTab
        myAppSendKeyRelease Key
        myAppSendKey vbKeyTab
    Else
        myAppSendKeyRelease Key
    End If

End Sub

'this function can be used to send input to the AC console (as it uses WM_CHAR)
Public Sub mySendKeyEx(Key As Long)
    Call myAppActivate
    Call mySendMessage(WM_KEYDOWN, Key, 0)
    Call mySendMessage(WM_CHAR, Key, 0)
    Call mySendMessage(WM_KEYUP, Key, 0)
End Sub

'Public Sub myGetWindowPos(ByVal hwnd As Long, ByRef lLeft As Long, ByRef lTop As Long, Optional ByRef lWidth As Long, Optional ByRef lHeight As Long)
'
'  Dim nWindowPlacement As WINDOWPLACEMENT
'
'  With nWindowPlacement
'    .Length = Len(nWindowPlacement)
'    Call GetWindowPlacement(hwnd, nWindowPlacement)
'
'    With .rcNormalPosition
'      lLeft = .Left '* Screen.TwipsPerPixelX
'      lTop = .Top ' * Screen.TwipsPerPixelY
'      lWidth = (.Right - .Left) '* Screen.TwipsPerPixelX
'      lHeight = (.Bottom - .Top) '* Screen.TwipsPerPixelY
'    End With
'
'  End With
'End Sub


Public Sub myMouseClick(ByVal xPos As Integer, ByVal yPos As Integer)
    Dim lParam As Long
    Dim wParam As Long

    'xPos = LOWORD(lParam);  // horizontal position of cursor
    'yPos = HIWORD(lParam);  // vertical position of cursor
    Call myAppActivate
    
    'Note: WM_LBUTTONDOWN/UP doesnt seem to work alone. I have to manually put
    'the mouse cursor above the place I want to be clicked, and then it will
    'click fine. That's why Im adding a MOUSEMOVE message before clicking
    
    lParam = (yPos * &H10000) + xPos
    
    MyDebug "myMouseClick: xPos = " & xPos & ", yPos = " & yPos & " -- lParam = " & Hex(lParam)
    
    'Doesn't move the mouse onscreen and works out of focus
    wParam = 0
    Call mySendMessage(WM_MOUSEMOVE, wParam, lParam)
    
    ' Doesn't work when not in Focus and moves mouse on screen (bleh)
    'Call MoveMouseCursor(xPos, yPos, g_PluginSite.hwnd)
    
    wParam = MK_LBUTTON
    Call mySendMessage(WM_LBUTTONDOWN, wParam, lParam)
    Call mySendMessage(WM_LBUTTONUP, wParam, lParam)
    
    'Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    'Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

End Sub

'Move mouse inside a window (or onscreen if window is zero)
Sub MoveMouseCursor(ByVal x As Long, ByVal y As Long, Optional ByVal hwnd As Long)
    If hwnd = 0 Then
        SetCursorPos x, y
    Else
        Dim lpPoint As POINTAPI
        lpPoint.x = x
        lpPoint.y = y
        Call ClientToScreen(hwnd, lpPoint)
        Call SetCursorPos(lpPoint.x, lpPoint.y)
    End If
End Sub


Public Function KeyIsDown(ByVal lKey As Long) As Boolean
    KeyIsDown = (GetKeyState(lKey) And &H8000)
End Function
