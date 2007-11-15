Attribute VB_Name = "WinAPI"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long 'elapsed time in msec since system boot
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long 'returns handle of the active window
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLength As Long)


'#define WM_MOUSEFIRST                   0x0200
'#define WM_MOUSEMOVE                    0x0200
'#define WM_LBUTTONDOWN                  0x0201
'#define WM_LBUTTONUP                    0x0202
'#define WM_LBUTTONDBLCLK                0x0203
'#define WM_RBUTTONDOWN                  0x0204
'#define WM_RBUTTONUP                    0x0205
'#define WM_RBUTTONDBLCLK                0x0206
'#define WM_MBUTTONDOWN                  0x0207
'#define WM_MBUTTONUP                    0x0208
'#define WM_MBUTTONDBLCLK                0x0209


Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
    
Public Const WM_ACTIVATE = &H6
Public Const WA_ACTIVATE = &H1

Public Const TabKey = &H9&
Public Const WM_PASTE = &H302

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Public Const MK_LBUTTON = &H1&
Public Const MK_RBUTTON = &H2&
Public Const MK_SHIFT = &H4&
Public Const MK_CONTROL = &H8&
Public Const MK_MBUTTON = &H10&

'----------------------

Public Function LoWord(ByRef DWord As Long) As Integer
    CopyMemory LoWord, ByVal VarPtr(DWord), 2
End Function

Public Function HiWord(ByRef DWord As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(DWord) + 2, 2
End Function

Public Function GET_X_LPARAM(ByVal lParam As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)
    GET_X_LPARAM = CLng("&H" & Right(hexstr, 4))
End Function

Public Function GET_Y_LPARAM(ByVal lParam As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(lParam), 8)
    GET_Y_LPARAM = CLng("&H" & Left(hexstr, 4))
End Function

Public Function GetElapsedSeconds() As Double
On Error GoTo ErrorHandler

    If (GetTickCount <= 0) Then
        GetElapsedSeconds = 1
    Else
        GetElapsedSeconds = CDbl(GetTickCount) / CDbl(1000)
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "GetElapsedSeconds - " & Err.Description
    GetElapsedSeconds = 1
    Resume Fin
End Function

Public Function GetACWindowRect() As RECT
On Error GoTo ErrorHandler

    Dim acRect As RECT
    Call GetWindowRect(g_PluginSite.hwnd, acRect)
    
Fin:
    GetACWindowRect = acRect
    Exit Function
ErrorHandler:
    PrintErrorMessage "GetACWindowRect - " & Err.Description
    Resume Fin
End Function



