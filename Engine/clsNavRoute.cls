VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNavRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Head As clsNavWaypoint    'linked list
Private m_Last As clsNavWaypoint
Private m_lCount As Long
Private m_db As DataFile
Private m_bLoaded As Boolean
Public Description As String 'route name

Private Const TAG_DATA_TYPE = "dat"
    Private Const DAT_TYPE_ROUTE = "route"
    Private Const DAT_TYPE_WP = "wp"
    
Private Const TAG_DESC = "desc" 'route or waypoint description

'Waypoint data
Private Const TAG_X = "x"
Private Const TAG_Y = "y"
Private Const TAG_Z = "z"
Private Const TAG_COORDS = "coord"


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    Description = "Default Route"
    Set m_Head = New clsNavWaypoint 'ghost
    Set m_Last = m_Head
    Set m_db = New DataFile
    m_bLoaded = False
    m_lCount = 0
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    Call ClearRoute
    Set m_Last = Nothing
    Set m_Head = Nothing
    Set m_db = Nothing
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavRoute.Class_Terminate - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get NumWP() As Long
    NumWP = m_lCount
End Property

Public Property Get FirstWP() As clsNavWaypoint
    Set FirstWP = m_Head.NextWp
End Property


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Private
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Public
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'Load a route from a file
Public Function LoadRoute(ByVal sFilePath As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim dat As clsDataEntry
    Dim wp As clsNavWaypoint
    
    MyDebug "Loading Route : " & sFilePath
    
    'make sure our route object is empty
    Call ClearRoute
    
    If m_db.Load(sFilePath) Then
        For Each dat In m_db
            If dat.ParamExist(TAG_DATA_TYPE) Then
                Dim sDataType As String: sDataType = dat.Param(TAG_DATA_TYPE)
                
                'route data
                If sDataType = DAT_TYPE_ROUTE Then
                    If dat.ParamExist(TAG_DESC) Then
                        Description = dat.Param(TAG_DESC)
                    Else
                        PrintWarning "LoadRoute : route has no name, using default"
                        Description = "Unnamed Route"
                    End If
                    
                'waypoint data
                ElseIf sDataType = DAT_TYPE_WP Then
                    If dat.ParamExist(TAG_DESC) And dat.ParamExist(TAG_X) And dat.ParamExist(TAG_Y) And dat.ParamExist(TAG_Z) Then
                        Set wp = New clsNavWaypoint
                        wp.Description = dat.Param(TAG_DESC)
                        wp.x = Val(dat.Param(TAG_X))
                        wp.y = Val(dat.Param(TAG_Y))
                        wp.z = Val(dat.Param(TAG_Z))
                        
                        If dat.ParamExist(TAG_COORDS) Then
                            wp.Coords = dat.Param(TAG_COORDS)
                        End If
                        
                        'Add wp to route
                        Call AddWaypoint(wp)
                        
                    Else
                        PrintWarning "LoadRoute : Waypoint Entry without Name or X/Y/Z - ignoring."
                    End If
                Else
                    MyDebug "LoadRoute : unhandled data type " & sDataType
                End If
            End If
        Next dat
        
        MyDebug "Route loaded - " & m_lCount & " waypoint(s)"
        m_bLoaded = True
        
    Else
        PrintErrorMessage "Failed to load Route : " & sFilePath & " - Reason : " & m_db.GetLastError
        m_bLoaded = False
    End If
    
Fin:
    Set dat = Nothing
    Set wp = Nothing
    LoadRoute = m_bLoaded
    Exit Function
ErrorHandler:
    m_bLoaded = False
    PrintErrorMessage "clsNavRoute.LoadRoute(" & sFilePath & ")"
    Resume Fin
End Function

'Save a route to a file
Public Function SaveRoute(Optional ByVal sFilePath As String = "")
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    Dim dat As clsDataEntry
    Dim wp As clsNavWaypoint
    
    If sFilePath = "" Then
        If m_db.FileName = "" Then
            PrintErrorMessage "SaveRoute() : route has no filename yet. Please specify one."
            bRet = False
            GoTo Fin
        Else
            sFilePath = m_db.FileName
        End If
    End If
    
    MyDebug "Saving Route '" & Description & "' to " & sFilePath
    
    'remove all the current entries from the database
    Call m_db.ResetData
    
    'Write route info
    Set dat = New clsDataEntry
    Call dat.AddParam(TAG_DATA_TYPE, DAT_TYPE_ROUTE)
    Call dat.AddParam(TAG_DESC, Description)
    
    If Not m_db.AddData(dat) Then
        PrintErrorMessage "SaveRoute(" & sFilePath & ") : failed to add the Route Info record to the database."
        bRet = False
        GoTo Fin
    End If
        
    'Write Waypoints Data
    Set wp = FirstWP
    While Valid(wp)
        Set dat = New clsDataEntry
        Call dat.AddParam(TAG_DATA_TYPE, DAT_TYPE_WP)
        Call dat.AddParam(TAG_DESC, wp.Description)
        Call dat.AddParam(TAG_X, wp.x)
        Call dat.AddParam(TAG_Y, wp.y)
        Call dat.AddParam(TAG_Z, wp.z)
        Call dat.AddParam(TAG_COORDS, wp.Coords)
        
        If Not m_db.AddData(dat) Then
            PrintErrorMessage "SaveRoute : failed to add waypoint " & wp.Description & " to data file."
            bRet = False
            GoTo Fin
        End If
        
        'move on to the next waypoint in the list
        Set wp = wp.NextWp
    Wend
    
    If m_db.save(sFilePath) Then
        m_db.FileName = sFilePath
        MyDebug "Route Saved to " & sFilePath & "."
        bRet = True
    Else
        PrintErrorMessage "SaveRoute(" & sFilePath & ") : save operation failed - " & m_db.GetLastError
        bRet = False
    End If
    
Fin:
    Set dat = Nothing
    Set wp = Nothing
    SaveRoute = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNavRoute.SaveRoute(" & sFilePath & ")"
    Resume Fin
End Function

'Removes all waypoints from the route
Public Sub ClearRoute()
On Error GoTo ErrorHandler
    
    Dim wp As clsNavWaypoint
    Dim tmpWp As clsNavWaypoint
    
    Set wp = m_Head
    While Valid(wp)
        Set tmpWp = wp.NextWp
        Set wp.NextWp = Nothing
        Set wp = tmpWp
    Wend

    Set m_Last = m_Head
    Set m_Head.NextWp = Nothing
    Set m_Head.PrevWp = Nothing
    Set m_Last.NextWp = Nothing
    Set m_Last.PrevWp = Nothing
    m_lCount = 0
    
Fin:
    Set wp = Nothing
    Set tmpWp = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavRoute.ClearRoute"
    Resume Fin
End Sub

'Inverts the waypoints order
Public Sub ReverseWaypoints()
On Error GoTo ErrorHandler
    
    Dim wp As clsNavWaypoint
    Dim NextWp As clsNavWaypoint
    Dim revWp As clsNavWaypoint     'current head of the reverse wp chain
    Dim PrevRevWp As clsNavWaypoint
    
    Set PrevRevWp = Nothing
    Set revWp = Nothing
    
    Set wp = m_Head.NextWp
    
    If Not Valid(wp) Then
        Set m_Last = m_Head
    Else
        Set m_Last = wp
        
        While Valid(wp)
            Set NextWp = wp.NextWp
            
            Set revWp = wp  'set the current head wp of the reverse chain
            Set revWp.NextWp = PrevRevWp    'set its next wp to the previous reverse wp chain head wp
            Set PrevRevWp = wp              'update reverse wp chain head
            
            Set wp = NextWp 'next waypoint in the regular waypoint chain
        Wend
        
        'set new head
        Set m_Head.NextWp = revWp
    End If
    
    Set m_Last.NextWp = Nothing
    
Fin:
    Set wp = Nothing
    Set NextWp = Nothing
    Set revWp = Nothing
    Set PrevRevWp = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavRoute.ReverseWaypoints"
    Resume Fin
End Sub

'Add a given WP to the route
Public Function AddWaypoint(wp As clsNavWaypoint) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    
    If Valid(wp) Then
        Set m_Last.NextWp = wp
        Set wp.NextWp = Nothing
        Set m_Last = wp
        m_lCount = m_lCount + 1
        bRet = True
    Else
        PrintErrorMessage "clsNavRoute.AddWaypoint : invalid waypoint"
        GoTo Fin
    End If
    
Fin:
    AddWaypoint = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNavRoute.AddWaypoint"
    Resume Fin
End Function

'Add a route WP at the current player location
Public Sub AddCurLoc(Optional ByVal sWPDesc As String = "")
On Error GoTo ErrorHandler

    Dim wp As New clsNavWaypoint

    wp.x = g_ds.AccuratePlayerLoc.Longitude
    wp.y = g_ds.AccuratePlayerLoc.Latitude
    wp.z = g_ds.AccuratePlayerLoc.Zoff
    wp.Coords = g_ds.AccuratePlayerLoc.Coords
    
    'wp.x = g_Hooks.LocationX
    'wp.y = g_Hooks.LocationY
    'wp.z = g_Hooks.LocationZ
    
    If Trim(sWPDesc) = "" Then
        wp.Description = "WP" & m_lCount
    Else
        wp.Description = sWPDesc
    End If
    
    If AddWaypoint(wp) Then
        PrintMessage "Waypoint added at (" & wp.x & ", " & wp.y & ", " & wp.z & ")"
    Else
        PrintErrorMessage "Couldnt add waypoint."
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavRoute.AddCurLoc"
    Resume Fin
End Sub


