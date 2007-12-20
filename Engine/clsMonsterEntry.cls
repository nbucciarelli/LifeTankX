VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonsterEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MonsterName As String
Public MonsterVuln As Integer
Public Yield As Boolean
Public Vuln As Boolean
Public Imperil As Boolean
Public Priority As Integer
Public Enabled As Boolean 'attack monster?

Private Sub Class_Initialize()
    MonsterName = "UnknownMonster"
    MonsterVuln = DMG_SLASHING
    Yield = False
    Vuln = False
    Imperil = False
    Priority = 1
    Enabled = True
End Sub
