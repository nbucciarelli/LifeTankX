VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLootFiltersManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'collections of clsLootFilter
Private m_colWeaponFilters As New Collection
Private m_colArmorFilters As New Collection
Private m_colWandFilters As New Collection
Private m_colSalvageFilters As New Collection
Private m_db As New DataFile

Private m_lLastWeapId As Long
Private m_lLastArmorId As Long
Private m_lLastWandId As Long
Private m_lLastSalvageId As Long

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    m_lLastWeapId = 0
    m_lLastArmorId = 0
    m_lLastWandId = 0
    m_lLastSalvageId = 0
End Sub

Private Sub Class_Terminate()
    Set m_colWeaponFilters = Nothing
    Set m_colArmorFilters = Nothing
    Set m_colWandFilters = Nothing
    Set m_colSalvageFilters = Nothing
    Set m_db = Nothing
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               PROPERTIES
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get WeaponFilters() As Collection
    Set WeaponFilters = m_colWeaponFilters
End Property

Public Property Get ArmorFilters() As Collection
    Set ArmorFilters = m_colArmorFilters
End Property

Public Property Get WandFilters() As Collection
    Set WandFilters = m_colWandFilters
End Property

Public Property Get SalvageFilters() As Collection
    Set SalvageFilters = m_colSalvageFilters
End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                                 PRIVATE
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Function GetNewId(ByRef lCounter As Long) As Long
On Error GoTo ErrorHandler
    Dim lRet As Long
    
    lRet = lCounter
    lCounter = lCounter + 1

Fin:
    GetNewId = lRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.GetNewId - " & Err.Description
    lRet = -1
    Resume Fin
End Function


Private Function AddFilter(colFilter As Collection, oFilter As clsLootFilter) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(oFilter) Then
        If Not FilterExist(oFilter.FilterId, colFilter) Then
            Call colFilter.Add(oFilter, CStr(oFilter.FilterId))
            bRet = True
        Else
            PrintErrorMessage "clsLootFilterManager.AddFilter - Filter #" & oFilter.FilterId & " is already in the collection."
            bRet = False
        End If
    End If

Fin:
    AddFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.AddFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           PUBLIC METHODS
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Function FilterExist(ByVal lFilterId As Long, colFilter As Collection, Optional ByRef oFilterOut As clsLootFilter) As Boolean
On Error GoTo NotFound
    Dim bRet As Boolean
    Set oFilterOut = colFilter(CStr(lFilterId))
    bRet = True
Fin:
    FilterExist = bRet
    Exit Function
NotFound:
    Set oFilterOut = Nothing
    bRet = False
    Resume Fin
End Function

'oFilter.FilterId will be modified
Public Function AddArmorFilter(ByRef oFilter As clsLootFilter) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(oFilter) Then
        If oFilter.FilterType = FILTER_ARMOR Then
            oFilter.FilterId = GetNewId(m_lLastArmorId)
            bRet = AddFilter(m_colArmorFilters, oFilter)
        Else
            PrintErrorMessage "clsLootFilterManager.AddArmorFilter - Tryed to add a non Armor filter to the Armor Filters collection"
        End If
    End If

Fin:
    AddArmorFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.AddArmorFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

'oFilter.FilterId will be modified
Public Function AddWeaponFilter(ByRef oFilter As clsLootFilter) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(oFilter) Then
        If oFilter.FilterType = FILTER_WEAPON Then
            oFilter.FilterId = GetNewId(m_lLastWeapId)
            bRet = AddFilter(m_colWeaponFilters, oFilter)
        Else
            PrintErrorMessage "clsLootFilterManager.AddWeaponFilter - Tryed to add a non Weapon filter to the Weapon Filters collection"
        End If
    End If

Fin:
    AddWeaponFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.AddWeaponFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

'oFilter.FilterId will be modified
Public Function AddWandFilter(ByRef oFilter As clsLootFilter) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(oFilter) Then
        If oFilter.FilterType = FILTER_WAND Then
            oFilter.FilterId = GetNewId(m_lLastWandId)
            bRet = AddFilter(m_colWandFilters, oFilter)
        Else
            PrintErrorMessage "clsLootFilterManager.AddWandFilter - Tryed to add a non Wand filter to the Wand Filters collection"
        End If
    End If

Fin:
    AddWandFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.AddWandFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

'oFilter.FilterId will be modified
Public Function AddSalvageFilter(ByRef oFilter As clsLootFilter) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If Valid(oFilter) Then
        If oFilter.FilterType = FILTER_SALVAGE Then
            oFilter.FilterId = GetNewId(m_lLastSalvageId)
            bRet = AddFilter(m_colSalvageFilters, oFilter)
        Else
            PrintErrorMessage "clsLootFilterManager.AddSalvageFilter - Tryed to add a non Salvage Filter to the Salvage Filters collection"
        End If
    End If

Fin:
    AddSalvageFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.AddSalvageFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function SalvageCombineFilter(colFilter As Collection, ByVal lFilterId As Long) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean

    If FilterExist(lFilterId, colFilter) Then
        Dim nFilter As clsLootFilter
        Set nFilter = colFilter(CStr(lFilterId))
        If (nFilter.SalvageCombine) Then
            nFilter.SalvageCombine = False
        Else
            nFilter.SalvageCombine = True
        End If
        bRet = True
    End If

Fin:
    SalvageCombineFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.CombineSalvageFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function EnableFilter(colFilter As Collection, ByVal lFilterId As Long) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If FilterExist(lFilterId, colFilter) Then
        Dim nFilter As clsLootFilter
        Set nFilter = colFilter(CStr(lFilterId))
        If (nFilter.FilterEnabled) Then
            nFilter.FilterEnabled = False
        Else
            nFilter.FilterEnabled = True
        End If
        bRet = True
    End If

Fin:
    EnableFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.EnableFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function


Public Function RemoveFilter(colFilter As Collection, ByVal lFilterId As Long) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    If FilterExist(lFilterId, colFilter) Then
        Call colFilter.Remove(CStr(lFilterId))
        bRet = True
    End If

Fin:
    RemoveFilter = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.RemoveFilter - " & Err.Description
    bRet = False
    Resume Fin
End Function

'Load filters from file
Public Function LoadFilters(ByVal sPath As String) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim dat As clsDataEntry
    Dim oFilter As clsLootFilter
    
    Call m_db.ResetData
    
    'if file doesnt exist, nothing to load, but don't return any error
    If Not PathExists(sPath) Then
        m_db.FileName = sPath
        bRet = True
        GoTo Fin
    End If
    
    If m_db.Load(sPath) Then
        For Each dat In m_db
            Set oFilter = New clsLootFilter
            
            If Not oFilter.LoadFromDataEntry(dat) Then
                PrintErrorMessage "clsLootFilterManager.LoadFilters - Unable to load a filter data"
            Else
                Select Case oFilter.FilterType
                    Case FILTER_ARMOR
                        If Not AddArmorFilter(oFilter) Then
                            PrintErrorMessage "clsLootFilterManager.LoadFilters - Failed to add Armor Filter #" & oFilter.FilterId
                        End If
                        
                    Case FILTER_WEAPON
                        If Not AddWeaponFilter(oFilter) Then
                            PrintErrorMessage "clsLootFilterManager.LoadFilters - Failed to add Weapon Filter #" & oFilter.FilterId
                        End If
                        
                    Case FILTER_WAND
                        If Not AddWandFilter(oFilter) Then
                            PrintErrorMessage "clsLootFilterManager.LoadFilters - Failed to add Wand Filter #" & oFilter.FilterId
                        End If
                        
                    Case FILTER_SALVAGE
                        If Not AddSalvageFilter(oFilter) Then
                            PrintErrorMessage "clsLootFilterManager.LoadFilters - Failed to add Salvage Filter #" & oFilter.FilterId
                        End If
                    
                    Case Else
                        MyDebug "WARNING - clsLootFilterManager.LoadFilters - Unsupported Filter Type " & oFilter.FilterType & " - Ignoring."
                End Select
            End If
            
            Set oFilter = Nothing
        Next dat
        MyDebug "Loot Filters loaded successfully from " & sPath
        bRet = True
    Else
        PrintErrorMessage "clsLootFilterManager.LoadFilters - Unable to load " & sPath & " : " & m_db.GetLastError
        bRet = False
    End If
    
Fin:
    Set dat = Nothing
    LoadFilters = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.LoadFilters(" & sPath & ") - " & Err.Description
    bRet = False
    Resume Fin
End Function

'Save filters to file
Public Function SaveFilters(Optional ByVal sPath As String = "") As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim lCount As Long
    Dim oFilter As clsLootFilter
    
    If sPath = "" Then sPath = m_db.FileName
    lCount = 0

    Call m_db.ResetData
    
    'Save armor filters
    For Each oFilter In m_colArmorFilters
        If Not oFilter.AddToDatabase(m_db) Then
            PrintErrorMessage "clsLootFilterManager.SaveFilters - Failed to add Armor Filter #" & oFilter.FilterId & " to database"
        Else
            lCount = lCount + 1
        End If
    Next oFilter
    
    'Save Weapon filters
    For Each oFilter In m_colWeaponFilters
        If Not oFilter.AddToDatabase(m_db) Then
            PrintErrorMessage "clsLootFilterManager.SaveFilters - Failed to add Weapon Filter #" & oFilter.FilterId & " to database"
        Else
            lCount = lCount + 1
        End If
    Next oFilter
    
    'Save Wand filters
    For Each oFilter In m_colWandFilters
        If Not oFilter.AddToDatabase(m_db) Then
            PrintErrorMessage "clsLootFilterManager.SaveFilters - Failed to add Wand Filter #" & oFilter.FilterId & " to database"
        Else
            lCount = lCount + 1
        End If
    Next oFilter
    
    'Save Salvage filters
    For Each oFilter In m_colSalvageFilters
        If Not oFilter.AddToDatabase(m_db) Then
            PrintErrorMessage "clsLootFilterManager.SaveFilters - Failed to add Salvage Filter #" & oFilter.FilterId & " to database"
        Else
            lCount = lCount + 1
        End If
    Next oFilter
    
    If m_db.save(sPath) Then
        MyDebug "Loot Filters (" & lCount & ") saved to " & sPath & " successfully"
        bRet = True
    Else
        PrintErrorMessage "clsLootFilterManager.SaveFilters - Failed to save Filters to " & sPath & " : " & m_db.GetLastError
        bRet = False
    End If
    
Fin:
    SaveFilters = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.SaveFilters(" & sPath & ") - " & Err.Description
    bRet = False
    Resume Fin
End Function

'Returns TRUE if the object passed at least one filter of the filters collection
Public Function PassFilters(objItem As acObject, colFilters As Collection) As Boolean
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim oFilter As clsLootFilter
    
    If Valid(objItem) And Valid(colFilters) Then
        For Each oFilter In colFilters
            If oFilter.PassFilter(objItem) Then
                bRet = True
                GoTo Fin
            End If
        Next oFilter
    Else
        PrintErrorMessage "clsLootFilterManager.PassFilters : invalid objItem or colFilters"
    End If
    
Fin:
    PassFilters = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsLootFilterManager.PassFilters - " & Err.Description
    bRet = False
    Resume Fin
End Function
