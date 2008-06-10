VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsACKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const KEY_ATTACK_LOW = "CombatLowAttack"
Private Const KEY_ATTACK_MEDIUM = "CombatMediumAttack"
Private Const KEY_ATTACK_HIGH = "CombatHighAttack"
Private Const KEY_JUMP = "MovementJump"
Private Const KEY_READY = "Ready"
Private Const KEY_CLOSESTMONSTER = "SelectionClosestMonster"


Public KeyAttackLow As Long
Public KeyAttackMedium As Long
Public KeyAttackHigh As Long
Public KeyJump As Long
Public KeyReady As Long
Public KeyClosestMonster As Long

Private Sub Class_Initialize()
    KeyAttackLow = -1      '211 DIK_DELETE
    KeyAttackMedium = -1   '207 DIK_END
    KeyAttackHigh = -1     '209 DIK_PGDN
    KeyReady = -1
End Sub

Public Sub LoadKeys()
On Error GoTo ErrorHandler
    
    MyDebug "--------- KeyBoard Settings ----------", True
    
    KeyAttackLow = g_PluginSite.QueryKeyboardMap(KEY_ATTACK_LOW)
    KeyAttackMedium = g_PluginSite.QueryKeyboardMap(KEY_ATTACK_MEDIUM)
    KeyAttackHigh = g_PluginSite.QueryKeyboardMap(KEY_ATTACK_HIGH)
    KeyJump = g_PluginSite.QueryKeyboardMap(KEY_JUMP)
    KeyReady = g_PluginSite.QueryKeyboardMap(KEY_READY)
    KeyClosestMonster = g_PluginSite.QueryKeyboardMap(KEY_CLOSESTMONSTER)
    
    MyDebug "Ready         : " & KeyReady, True
    MyDebug "Jump          : " & KeyJump, True
    MyDebug "Attack Low : " & KeyAttackLow, True
    MyDebug "       Medium : " & KeyAttackMedium, True
    MyDebug "       High : " & KeyAttackHigh, True

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACKeys.LoadKeys - " & Err.Description & " line " & Erl
    Resume Fin
End Sub

Private Sub CheckKey(ByVal lKey As Long, ByVal sDesc As String)
On Error GoTo ErrorHandler
    
    If (lKey = -1) Or (lKey = 0) Then
        PrintErrorMessage "Warning - The key for " & sDesc & " could not be found / is not set. Please update your AC keyboard settings."
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACKeys.CheckKey(" & lKey & ", " & sDesc & ")"
    Resume Fin
End Sub

Public Sub PrintNotFound()
On Error GoTo ErrorHandler
    
    CheckKey KeyAttackLow, KEY_ATTACK_LOW
    CheckKey KeyAttackMedium, KEY_ATTACK_MEDIUM
    CheckKey KeyAttackHigh, KEY_ATTACK_HIGH
    CheckKey KeyReady, KEY_READY
    CheckKey KeyClosestMonster, KEY_CLOSESTMONSTER
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsACKeys.PrintNotFound"
    Resume Fin
End Sub
