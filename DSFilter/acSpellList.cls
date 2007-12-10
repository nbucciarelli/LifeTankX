VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acSpellList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Description As String

Private m_colFamily As Collection   'collection of spell family
                                    'NOTE: spells who don't belong to any family are their
                                    'own family
                                    
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Set m_colFamily = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colFamily = Nothing
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                   Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get Families() As Collection
    Set Families = m_colFamily
End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Family Collection Handlers
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Function CreateFamily(ByVal sFamilyName As String, ByVal iType As Integer, ByVal iElement As Integer, ByVal iSchool As Integer, ByVal lIcon As Long) As clsSpellFamily
On Error GoTo ErrorHandler
    
    Dim objNewFam As New clsSpellFamily
    With objNewFam
        .mFamily = sFamilyName
        .mType = iType
        .mElement = iElement
        .mSchool = iSchool
        .mIcon = lIcon
    End With

    Call m_colFamily.Add(objNewFam, objNewFam.mFamily)
    Set CreateFamily = objNewFam
    
    'MyDebug Description & " - Created Spell Family : " & CreateFamily.mFamily
    
Fin:
    Set objNewFam = Nothing
    Exit Function
ErrorHandler:
    Set CreateFamily = Nothing
    PrintErrorMessage "clsSpellList.CreateFamily"
    Resume Fin
End Function

'Default Class Method
Public Function GetFamily(ByVal sFamilyName As String) As clsSpellFamily
Attribute GetFamily.VB_UserMemId = 0
On Error GoTo NotFound
    Set GetFamily = m_colFamily(sFamilyName)
Fin:
    Exit Function
NotFound:
    Set GetFamily = Nothing
    Resume Fin
End Function

Private Function FindFamily(ByVal sFamilyName As String, ByRef objFamilyOut As clsSpellFamily) As Boolean
On Error GoTo ErrorHandler

    Set objFamilyOut = GetFamily(sFamilyName)
    FindFamily = Valid(objFamilyOut)
    
Fin:
    Exit Function
ErrorHandler:
    FindFamily = False
    PrintErrorMessage Description & ".FindFamily"
    Resume Fin
End Function

'AddSpell : puts the spell in the proper family
Private Sub AddSpell(objSpell As clsSpell)
On Error GoTo ErrorHandler

    Dim objFamily As clsSpellFamily

    If Not Valid(objSpell) Then
        PrintWarning Description & ".AddSpell : Invalid objSpell"
        Exit Sub
    End If
    
    If objSpell.HasFamily And FindFamily(objSpell.SpellFamily, objFamily) Then
        Call objFamily.AddSpell(objSpell)
    Else
        'if this spell has no family, then make sure its familyname = spellname
        If Not objSpell.HasFamily Then
            objSpell.SpellFamily = objSpell.SpellName
        End If
        
        Set objFamily = CreateFamily(objSpell.SpellFamily, objSpell.SpellType, objSpell.SpellElement, objSpell.SpellSchool, objSpell.SpellIcon)
        If Valid(objFamily) Then
            Call objFamily.AddSpell(objSpell)
        Else
            PrintWarning Description & ".AddSpell : error creating family"
        End If
    End If


Fin:
    Set objFamily = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage Description & ".AddSpell"
    Resume Fin
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Spell Finding Methods
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Function FindSpellInFamily(objFamily As clsSpellFamily, Optional iWantedLevel As Integer = 7, Optional sPartialSpellName As String = "", Optional bIgnoreSkillReq As Boolean = False) As clsSpell
On Error GoTo ErrorHandler

    If Not Valid(objFamily) Then
        PrintWarning Description & ".FindSpellInFamily - invalid objFamily"
        Set FindSpellInFamily = Nothing
    Else
        Dim objSpell As clsSpell
        Dim bMatchName As Boolean
        Dim iLevel As Integer
        Dim iAllowedLevel As Integer
        Dim iHighestLvl As Integer
        Dim objSelectedSpell As clsSpell
        
        bMatchName = (sPartialSpellName <> "")
        
        'make sure desired spell level is valid/available
        If (iWantedLevel < 1) Then
            iWantedLevel = 1
        ElseIf iWantedLevel > objFamily.HighestSpellLevel Then
            iWantedLevel = objFamily.HighestSpellLevel
        End If
    
        iLevel = iWantedLevel
        
        If Not bIgnoreSkillReq Then
            iAllowedLevel = SkillToSpellLevel(g_Hooks.Skill(MagicSchoolToSkillId(objFamily.mSchool)))
            If iWantedLevel > iAllowedLevel Then iLevel = iAllowedLevel
        End If
    
        iHighestLvl = 0
        Set objSelectedSpell = Nothing
        For Each objSpell In objFamily
            If (objSpell.SpellLevel <= iLevel) And (objSpell.SpellLevel >= iHighestLvl) Then
                If ((bMatchName And InStr(LCase(objSpell.SpellName), LCase(sPartialSpellName))) Or (Not bMatchName)) Then
                    If g_Filters.SpellLearned(objSpell.SpellID) Then
                        iHighestLvl = objSpell.SpellLevel
                        Set objSelectedSpell = objSpell
                    End If
                End If
            End If
        Next objSpell
        
        Set FindSpellInFamily = objSelectedSpell
    End If
    
Fin:
    Exit Function
ErrorHandler:
    Set FindSpellInFamily = Nothing
    PrintErrorMessage Description & ".FindSpellInFamily"
    Resume Fin
End Function

Public Function FindSpell(ByVal sFamilyName As String, Optional iWantedLevel As Integer = 7, Optional sSpellName As String = "", Optional bIgnoreSkillReq As Boolean = False) As clsSpell
On Error GoTo ErrorHandler

    Dim objFamily As clsSpellFamily
    
    'MyDebug "Trying to find spell of family : " & sFamilyName & ", level " & iWantedLevel
    
    If FindFamily(sFamilyName, objFamily) Then
        Set FindSpell = FindSpellInFamily(objFamily, iWantedLevel, sSpellName, bIgnoreSkillReq)
    Else
        PrintWarning Description & ".FindSpell : couldn't find spell family " & sFamilyName
        Set FindSpell = Nothing
    End If
    
Fin:
    Set objFamily = Nothing
    Exit Function
ErrorHandler:
    Set FindSpell = Nothing
    PrintErrorMessage Description & ".FindSpell"
    Resume Fin
End Function

Public Function FindSpellByID(ByVal lSpellID As Long) As clsSpell
On Error GoTo ErrorHandler

    Dim objFamily As clsSpellFamily
    Dim objSpell As clsSpell
        
    For Each objFamily In m_colFamily
        For Each objSpell In objFamily
            If objSpell.SpellID = lSpellID Then
                GoTo Fin
            End If
        Next objSpell
    Next objFamily
    
    'not found
    Set objSpell = Nothing
    
Fin:
    Set FindSpellByID = objSpell
    Set objSpell = Nothing
    Set objFamily = Nothing
    Exit Function
ErrorHandler:
    Set FindSpellByID = Nothing
    PrintErrorMessage Description & ".FindSpellByID"
    Resume Fin
End Function

Public Function FindSpellByType(ByVal iType As Integer, ByVal iElement As Integer, Optional iWantedLevel As Integer = 7, Optional bIgnoreSkillReq As Boolean = False) As clsSpell
On Error GoTo ErrorHandler
    
    Dim objFamily As clsSpellFamily
    
    'MyDebug "Trying to find spell by type : " & GetSpelltypeString(iType) & ", with element " & GetDamageString(iElement)
    
    For Each objFamily In m_colFamily
        If (objFamily.mType = iType) And (objFamily.mElement = iElement) Then
            'MyDebug "FindSpellByType - Found potential family : " & objFamily.mFamily
            Set FindSpellByType = FindSpellInFamily(objFamily, iWantedLevel, , bIgnoreSkillReq)
            Exit Function
        End If
    Next objFamily
    
    'not found
    Set FindSpellByType = Nothing
    
Fin:
    Set objFamily = Nothing
    Exit Function
ErrorHandler:
    Set FindSpellByType = Nothing
    PrintErrorMessage Description & ".FindSpellByType"
    Resume Fin
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Spells Database Loading
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function LoadData(ByVal sFileName As String) As Boolean
On Error GoTo Error_Handler

    Dim bRet As Boolean
    Dim db As New DataFile
    Dim dat As clsDataEntry
    Dim sPath As String
    Dim objSpell As clsSpell

    sPath = App.Path & "\" & sFileName
    
    If Not db.Load(sPath) Then
        PrintErrorMessage Description & ".LoadData : failed to load " & sPath
        LoadData = False
        Exit Function
    End If
    
    For Each dat In db
        Set objSpell = New clsSpell
        With objSpell
            If dat.ParamExist(TAG_SPELL_NAME) Then .SpellName = dat.Param(TAG_SPELL_NAME)
            If dat.ParamExist(TAG_SPELL_TYPE) Then .SpellType = dat.Param(TAG_SPELL_TYPE)
            If dat.ParamExist(TAG_SPELL_FAMILY) Then .SpellFamily = dat.Param(TAG_SPELL_FAMILY)
            If dat.ParamExist(TAG_SPELL_SCHOOL) Then .SpellSchool = dat.Param(TAG_SPELL_SCHOOL)
            If dat.ParamExist(TAG_SPELL_LEVEL) Then .SpellLevel = dat.Param(TAG_SPELL_LEVEL)
            If dat.ParamExist(TAG_SPELL_ELEMENT) Then .SpellElement = dat.Param(TAG_SPELL_ELEMENT)
            If dat.ParamExist(TAG_SPELL_ID) Then .SpellID = dat.Param(TAG_SPELL_ID)
            If dat.ParamExist(TAG_SPELL_ICON) Then .SpellIcon = dat.Param(TAG_SPELL_ICON)
        End With
        
        'Add the spell to the list
        Call AddSpell(objSpell)
        Set objSpell = Nothing
    Next dat
    
    bRet = True
Fin:
    LoadData = bRet
    Exit Function
Error_Handler:
    bRet = False
    PrintErrorMessage Description & ".LoadData  : error while loading " & sPath
    Resume Fin
End Function


Public Sub Display()
    
End Sub
