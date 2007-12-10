VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SpellName As String
Public SpellID As Long
Public SpellLevel As Integer
Public SpellType As Integer
Public SpellFamily As String
Public SpellIcon As Long
Public SpellElement As Integer
Public SpellSchool As Integer

Private Sub Class_Initialize()
    SpellFamily = NO_FAMILY
    SpellType = SPELLTYPE_NORMAL
    SpellLevel = 0
    SpellName = "UnknownSpell"
    SpellID = 0
    SpellIcon = 0
    SpellSchool = SCHOOL_CREATURE
    SpellElement = DMG_NONE
End Sub


Public Function HasFamily() As Boolean
    HasFamily = (SpellFamily <> NO_FAMILY) And (SpellName <> SpellFamily)
End Function
