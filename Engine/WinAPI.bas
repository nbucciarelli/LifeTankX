Attribute VB_Name = "shWinAPI"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Public Declare Function timeGetTime Lib "kernel32" () As Long 'elapsed time in msec since system boot

'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Public Declare Function GetActiveWindow Lib "user32" () As Long 'returns handle of the active window
'Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As rect) As Long

'Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal numBytesToRead As Long)
    
'Public Declare Sub CopyMemToStr Lib "kernel32" Alias "RtlMoveMemory" _
'  (ByVal sDest As String, Source As Any, ByVal numBytesToRead As Long)

Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long


Public Type rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

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

Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As rect
End Type

Public Const MK_LBUTTON = &H1&
Public Const MK_RBUTTON = &H2&
Public Const MK_SHIFT = &H4&
Public Const MK_CONTROL = &H8&
Public Const MK_MBUTTON = &H10&

Public Function ReadDWORD(dat() As Byte, ByVal dwOffset As Long) As Long
On Error GoTo ErrorHandler

    Dim dwRet As Long
    Call CopyMemory(dwRet, dat(dwOffset), 4)
    
Fin:
    ReadDWORD = dwRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "ReadDWORD - " & Err.Description
    dwRet = 0
    Resume Fin
End Function


Public Function GetACWindowRect() As rect
On Error GoTo ErrorHandler

    Dim acRect As rect
    Call GetWindowRect(g_PluginSite.hWnd, acRect)
    
Fin:
    GetACWindowRect = acRect
    Exit Function
ErrorHandler:
    PrintErrorMessage "GetACWindowRect - " & Err.Description
    Resume Fin
End Function


Public Function MouseX() As Long
    Dim lpPoint As POINTAPI
    Call ClientToScreen(g_PluginSite.hWnd, lpPoint)
    MouseX = lpPoint.x
End Function

Public Function MouseY() As Long
    Dim lpPoint As POINTAPI
    Call ClientToScreen(g_PluginSite.hWnd, lpPoint)
    MouseY = lpPoint.y
End Function
