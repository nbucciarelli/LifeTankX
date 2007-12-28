VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpellFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colSpells As Collection

Public mFamily As String
Public mSchool As Integer
Public mElement As Integer
Public mType As Integer
Public mIcon As Long
 
Private m_iHighestLevel As Integer   'highest spell level in the collection

Private Sub Class_Initialize()
    m_iHighestLevel = 1
    mFamily = NO_FAMILY
    mElement = DMG_NONE
    mIcon = 0
    mSchool = SCHOOL_CREATURE
    mType = SPELLTYPE_NORMAL
    Set m_colSpells = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colSpells = Nothing
End Sub

Public Property Get Members() As Collection
    Set Members = m_colSpells
End Property

Public Property Get HighestSpellLevel() As Integer
    HighestSpellLevel = m_iHighestLevel
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_colSpells.[_NewEnum]
End Function

Public Function FindSpellByLevel(ByVal iLevel As Integer, Optional ByRef objSpellOut As clsSpell) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    Dim objSpell As clsSpell
    
    For Each objSpell In m_colSpells
        If objSpell.SpellLevel = iLevel Then
            Set objSpellOut = objSpell
            bRet = True
            GoTo Fin
        End If
    Next objSpell
    
    bRet = False
    
Fin:
    FindSpellByLevel = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsFamily.FindSpellByLevel - " & Err.Description
    Resume Fin
End Function

'Default Class Method
Public Function GetSpell(ByVal sSpellName) As clsSpell
Attribute GetSpell.VB_UserMemId = 0
On Error GoTo NotFound
    Set GetSpell = m_colSpells(sSpellName)
Fin:
    Exit Function
NotFound:
    Set GetSpell = Nothing
    Resume Fin
End Function

Public Function FamilySpellByName(ByVal sSpellName As String, Optional ByRef objSpellOut As clsSpell) As Boolean
On Error GoTo ErrorHandler

    Set objSpellOut = GetSpell(sSpellName)
    FamilySpellByName = Valid(objSpellOut)
    
Fin:
    Exit Function
ErrorHandler:
    FamilySpellByName = False
    PrintErrorMessage "clsFamily.FamilySpellByName - " & Err.Description
    Resume Fin
End Function

Public Function AddSpell(objSpell As clsSpell) As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean

    'works for both spells with or without family

    If Not Valid(objSpell) Then
        PrintWarning "SpellFamily.AddSpell : invalid objSpell"
        bRet = False
    ElseIf objSpell.SpellFamily <> mFamily Then
        PrintWarning "SpellFamily.AddSpell : spell family " & objSpell.SpellFamily & " is different from " & mFamily
        bRet = False
    ElseIf objSpell.SpellType <> mType Then
        PrintWarning "SpellFamily.AddSpell : spell type " & GetSpelltypeString(objSpell.SpellType) & " is different from " & GetSpelltypeString(mType)
        bRet = False
    ElseIf objSpell.SpellElement <> mElement Then
        PrintWarning "SpellFamily.AddSpell : spell element " & GetDamageString(objSpell.SpellElement) & " is different from " & GetDamageString(mElement)
        bRet = False
    ElseIf objSpell.SpellSchool <> mSchool Then
        PrintWarning "SpellFamily.AddSpell : spell school " & GetSchoolString(objSpell.SpellSchool) & " is different from " & GetSchoolString(mSchool)
        bRet = False
    ElseIf FamilySpellByName(objSpell.SpellName) Then
        PrintWarning "SpellFamily.AddSpell : spell " & objSpell.SpellName & " is already in the collection"
        bRet = False
    Else
        If objSpell.SpellLevel > m_iHighestLevel Then
            m_iHighestLevel = objSpell.SpellLevel
            If objSpell.SpellIcon <> 0 Then mIcon = objSpell.SpellIcon  'update icon too
        End If
        
        'MyDebug "[" & mFamily & "] Added : " & objSpell.SpellName
        
        Call m_colSpells.Add(objSpell, objSpell.SpellName)
        bRet = True
    End If
    
Fin:
    AddSpell = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSpellFamily.AddSpell - " & Err.Description
    bRet = False
    Resume Fin
End Function
