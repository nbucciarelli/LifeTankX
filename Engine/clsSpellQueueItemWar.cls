VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellQueueItemSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Used by the SpellQueue class
Option Explicit

Public SpellID As Long
Public TargetGUID As Long
Public SpellName As String

Private Sub Class_Initialize()
    SpellID = -1
    TargetGUID = -1
    SpellName = "War Spell"
End Sub

Public Function Description() As String
    'Description = "War Spell"
    Description = SpellName
End Function
