VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellQueueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Used by the SpellQueue class
Option Explicit

Public Index As Long

Public SpellSchool As Integer
Public SpellFamily As String
Public SpellID As Long
Public SpellType As Integer
Public SpellElement As Integer
Public LevelWanted As Integer
Public TimeRemaining As Long
Public TargetGUID As Long
Public TargetName As String

Private Sub Class_Initialize()
    Index = -1
    SpellFamily = NO_FAMILY
    SpellType = SPELLTYPE_NORMAL
    SpellElement = DMG_NONE
    TargetGUID = 0
    TargetName = "NoTarget"
    LevelWanted = 8
    SpellSchool = SCHOOL_CREATURE
    TimeRemaining = 0
End Sub

Public Function Description() As String
    'If SpellType = SPELLTYPE_NORMAL Then
        Description = SpellFamily
    'Else
    '    Description = GetDamageString(SpellElement) & " " & GetSpelltypeString(SpellType)
    'End If
End Function
