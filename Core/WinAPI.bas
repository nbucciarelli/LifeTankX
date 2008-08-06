Attribute VB_Name = "WinAPI"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Public Declare Function timeGetTime Lib "kernel32" () As Long 'elapsed time in msec since system boot
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
 
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long 'returns handle of the active window
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLength As Long)

Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long


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

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MOVE = &H1

Public Const MK_LBUTTON = &H1&
Public Const MK_RBUTTON = &H2&
Public Const MK_SHIFT = &H4&
Public Const MK_CONTROL = &H8&
Public Const MK_MBUTTON = &H10&

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102

Public Const WM_SETFOCUS = &H7

Public Const WM_ACTIVATE = &H6
Public Const WA_ACTIVATE = &H1

Public Const TabKey = &H9&
Public Const WM_PASTE = &H302

Public Const BM_CLICK = &HF5
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type rect
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
  rcNormalPosition As rect
End Type


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

Public Function GetACWindowRect() As rect
On Error GoTo ErrorHandler

    Dim acRect As rect
    Call GetWindowRect(g_PluginSite.hwnd, acRect)
    
Fin:
    GetACWindowRect = acRect
    Exit Function
ErrorHandler:
    PrintErrorMessage "GetACWindowRect - " & Err.Description
    Resume Fin
End Function



