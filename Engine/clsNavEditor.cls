VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNavEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_NAV_INTERFACE = PATH_DATA & "\nav.xml"

Private Enum eWPListCols
    COL_WP_NAME
    COL_WP_COORDS
    COL_WP_REMOVE
End Enum

'View
Private m_NavView As DecalPlugins.IView
Private m_sViewShema As String
Private m_colRouteFiles As New colFileDir

'Controls
Private WithEvents chRoute As DecalControls.Choice
Attribute chRoute.VB_VarHelpID = -1
Private WithEvents chNavType As DecalControls.Choice
Attribute chNavType.VB_VarHelpID = -1
Private WithEvents cmdNewRoute As DecalControls.PushButton
Attribute cmdNewRoute.VB_VarHelpID = -1
Private WithEvents cmdRouteGo As DecalControls.PushButton
Attribute cmdRouteGo.VB_VarHelpID = -1
Private WithEvents cmdNavStop As DecalControls.PushButton
Attribute cmdNavStop.VB_VarHelpID = -1
Private WithEvents cmdAddWp As DecalControls.PushButton
Attribute cmdAddWp.VB_VarHelpID = -1
Private WithEvents btnWpRemove As DecalControls.PushButton
Attribute btnWpRemove.VB_VarHelpID = -1
Private WithEvents cmdGoToWp As DecalControls.PushButton
Attribute cmdGoToWp.VB_VarHelpID = -1
Private WithEvents cmdClear As DecalControls.PushButton
Attribute cmdClear.VB_VarHelpID = -1
Private WithEvents cmdCloseNav As DecalControls.PushButton
Attribute cmdCloseNav.VB_VarHelpID = -1
Private WithEvents cmdSaveRoute As DecalControls.PushButton
Attribute cmdSaveRoute.VB_VarHelpID = -1
Private WithEvents cmdSaveRouteAs As DecalControls.PushButton
Attribute cmdSaveRouteAs.VB_VarHelpID = -1
Private txtSaveAs As DecalControls.Edit
Private txtNewRoute As DecalControls.Edit
Private txtWp As DecalControls.Edit
Private WithEvents lstWp As DecalControls.list
Attribute lstWp.VB_VarHelpID = -1


'Vars
Private m_sCurRoute As String   'current route name, empty if no routes created/loaded

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
    m_sViewShema = FileToString(FILE_NAV_INTERFACE)
    m_sCurRoute = ""
    Set m_colRouteFiles = New colFileDir
    Call ScanRoutesFolder
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    Set chRoute = Nothing
    Set cmdNewRoute = Nothing
    Set cmdAddWp = Nothing
    Set btnWpRemove = Nothing
    Set cmdGoToWp = Nothing
    Set cmdClear = Nothing
    Set cmdRouteGo = Nothing
    Set cmdCloseNav = Nothing
    Set cmdSaveRoute = Nothing
    Set txtNewRoute = Nothing
    Set txtWp = Nothing
    Set lstWp = Nothing
    Set cmdSaveRouteAs = Nothing
    Set txtSaveAs = Nothing
    Set cmdNavStop = Nothing
    Set chNavType = Nothing

    Set m_NavView = Nothing
    Set m_colRouteFiles = Nothing

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor_Terminate - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get RouteFiles() As colFileDir
    Set RouteFiles = m_colRouteFiles
End Property

Public Property Get Opened() As Boolean
    If Valid(m_NavView) Then
        Opened = m_NavView.Activated
    Else
        Opened = False
    End If
End Property

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Private
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub chRoute_Change(ByVal nID As Long, ByVal nIndex As Long)
    If nIndex >= 0 Then
        Dim sRouteName As String
        Dim sRouteFilePath As String
        sRouteName = chRoute.Text(nIndex)
        
        sRouteFilePath = GetRouteFilePath(sRouteName)
        If Not FileExists(sRouteFilePath) Then
            PrintMessage "This route doesnt exists anymore - Updating list now."
            Call UpdateRoutesList
        Else
            If g_Nav.Route.LoadRoute(sRouteFilePath) Then
                m_sCurRoute = sRouteName
                Call UpdateWaypointsList
                Call UpdateRouteInfo
            Else
                PrintErrorMessage "Could not load route '" & sRouteName & "' from " & sRouteFilePath
            End If
        End If
    End If
End Sub

Private Sub cmdAddWp_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    Call g_Nav.Route.AddCurLoc(Trim(txtWp.Text))
    Call UpdateWaypointsList
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdAddWp_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub UpdateWaypointsList()
On Error GoTo ErrorHandler
    
    Dim wp As clsNavWaypoint
    Dim i As Integer
    
    Call lstWp.Clear
    
    Set wp = g_Nav.Route.FirstWP
    While Valid(wp)
        i = lstWp.AddRow
        lstWp.Data(COL_WP_NAME, i) = wp.Description
        lstWp.Data(COL_WP_COORDS, i) = wp.Coords
        Set wp = wp.NextWp
    Wend

Fin:
    Set wp = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdAddWp_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdClear_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    Call g_Nav.Route.ClearRoute
    Call UpdateWaypointsList
    PrintMessage "Route cleared."
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdClear_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdCloseNav_Accepted(ByVal nID As Long)
    Call HideEditor
End Sub

Private Sub cmdNewRoute_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler
    
    Dim sVal As String
    
    sVal = Trim(txtNewRoute.Text)
    
    If sVal <> "" Then
        Call CreateNewRoute(sVal)
        txtNewRoute.Text = ""
    Else
        PrintMessage "Please enter a valid route name."
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdNewRoute_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub CreateNewRoute(ByVal sRouteName As String)
On Error GoTo ErrorHandler

    sRouteName = Trim(sRouteName)
    m_sCurRoute = sRouteName
    Call g_Nav.MakeNewRoute(m_sCurRoute)
    Call g_Nav.Route.SaveRoute(GetRouteFilePath(sRouteName))
    Call UpdateRoutesList
    Call UpdateWaypointsList
    Call UpdateRouteInfo
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.CreateNewRoute - " & Err.Description
    Resume Fin
End Sub

Private Sub cmdNavStop_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    Call g_Nav.NavStop

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdRouteGo_Accepted - " & Err.Description
    Resume Fin
End Sub



Private Sub cmdRouteGo_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    Call g_Nav.ResumeRoute(False)

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdRouteGo_Accepted - " & Err.Description
    Resume Fin
End Sub


Private Sub cmdSaveRoute_Accepted(ByVal nID As Long)
On Error GoTo ErrorHandler

    If m_sCurRoute = "" Then
        PrintMessage "You must create a new route first."
    Else
        Call g_Nav.Route.SaveRoute
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.cmdSaveRoute_Accepted - " & Err.Description
    Resume Fin
End Sub

Private Sub chNavType_Change(ByVal nID As Long, ByVal nIndex As Long)
On Error GoTo ErrorHandler

    If nIndex >= 0 Then
        g_Nav.NavType = nIndex + 1
        MyDebug "chNavType_Change - Setting NavType to " & g_Nav.NavType & " (Index:" & nIndex & ")"
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.chNavType_Change - " & Err.Description
    Resume Fin
End Sub


Private Sub cmdSaveRouteAs_Accepted(ByVal nID As Long)
    Dim sVal As String
    
    sVal = Trim(txtSaveAs.Text)
    
    If sVal <> "" Then
        Call SaveRouteAs(sVal)
        txtSaveAs.Text = ""
    Else
        PrintMessage "Please enter a valid route name to save as."
    End If
    
End Sub

Private Sub UpdateRouteInfo()
On Error GoTo ErrorHandler

    If g_Nav.NavType > NAVTYPE_NONE And g_Nav.NavType < NAVTYPE_FOLLOW Then
        Dim lIndex As Long
        lIndex = g_Nav.NavType - 1

        If lIndex >= 0 And lIndex < chNavType.ChoiceCount Then
            chNavType.Selected = lIndex
        Else
            PrintErrorMessage "Base chNavType index " & lIndex
        End If
    Else
        If g_Nav.NavType = NAVTYPE_NONE Then    'force to loop
            g_Nav.NavType = NAVTYPE_LOOP
            Call UpdateRouteInfo
        End If
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.UpdateRouteInfo - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Public
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


Public Function LoadEditorView() As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    
    Set m_NavView = g_PluginSite.LoadView(m_sViewShema)

    If Valid(m_NavView) Then
        Set chRoute = m_NavView.Control("chRoute")
        Set cmdNewRoute = m_NavView.Control("cmdNewRoute")
        Set cmdAddWp = m_NavView.Control("cmdAddWp")
        Set btnWpRemove = m_NavView.Control("btnWpRemove")
        Set cmdGoToWp = m_NavView.Control("cmdGoToWp")
        Set cmdClear = m_NavView.Control("cmdClear")
        Set cmdRouteGo = m_NavView.Control("cmdRouteGo")
        Set cmdCloseNav = m_NavView.Control("cmdCloseNav")
        Set cmdSaveRoute = m_NavView.Control("cmdSaveRoute")
        Set txtNewRoute = m_NavView.Control("txtNewRoute")
        Set txtWp = m_NavView.Control("txtWp")
        Set lstWp = m_NavView.Control("lstWp")
        Set cmdSaveRouteAs = m_NavView.Control("cmdSaveRouteAs")
        Set txtSaveAs = m_NavView.Control("txtSaveAs")
        Set cmdNavStop = m_NavView.Control("cmdNavStop")
        Set chNavType = m_NavView.Control("chNavType")
        
        txtSaveAs.Text = ""
        txtWp.Text = ""
        txtNewRoute.Text = ""
        
        chNavType.Selected = 1  'Loop mode
        
        bRet = True
    Else
        PrintErrorMessage "clsNavEditor.LoadEditorView - bad m_NavView"
        bRet = False
    End If
    
    Call UpdateRoutesList
    Call UpdateWaypointsList
    
Fin:
    LoadEditorView = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNavEditor.LoadEditorView - " & Err.Description
    Resume Fin
End Function

Public Sub ShowEditor(Optional bAutoPosition As Boolean = True)
On Error GoTo ErrorHandler

    If m_NavView Is Nothing Then
        Call LoadEditorView
        
        If bAutoPosition Then
            Dim newPos As Decal.tagRECT
        
            With m_NavView.Position
                newPos.Bottom = .Bottom
                newPos.Left = .Left
                newPos.Right = .Right
                newPos.Top = .Top
            End With
        
            'newPos.Left = g_Hooks.Area3DWidth - newPos.Right - 50
            'newPos.Top = g_Hooks.Area3DHeight - newPos.Bottom - 50
            newPos.Left = g_Hooks.AC3DRegionRect.Right - g_Hooks.AC3DRegionRect.Left - newPos.Right - 50
            newPos.Top = g_Hooks.AC3DRegionRect.Top - g_Hooks.AC3DRegionRect.Bottom - newPos.Bottom - 50
        
            'update screen pos
            'm_NavView.Position = newPos
            m_NavView.Position.Top = newPos.Top
            m_NavView.Position.Bottom = newPos.Bottom
            m_NavView.Position.Left = newPos.Left
            m_NavView.Position.Right = newPos.Right
            'm_NavView.Position = newPos

        End If
    End If
    
    If Valid(m_NavView) Then
        Call m_NavView.Activate
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.ShowEditor - " & Err.Description
    Resume Fin
End Sub

Public Sub HideEditor()
On Error GoTo ErrorHandler

    If Valid(m_NavView) Then
        Call m_NavView.Deactivate
    End If
    
    'Set m_NavView = Nothing
    'Call g_PluginSite.RedrawBar
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNavEditor.ShowEditor - " & Err.Description
    Resume Fin
End Sub

'Get a list of the route files available in the Routes folder
Public Function ScanRoutesFolder() As Boolean
On Error GoTo ErrorHandler
    
    ScanRoutesFolder = m_colRouteFiles.DirFolder(g_Settings.GetDataFolder & "\" & PATH_ROUTES, vbNormal, "*." & FILE_EXT_ROUTE)

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsNavEditor.ScanRoutesFolder - " & Err.Description
    Resume Fin
End Function

Public Sub UpdateRoutesList()
On Error GoTo ErrorHandler

    Dim objFile As clsFileInfo
    
    MyDebug "UpdateRoutesList"
    
    Call chRoute.Clear
    If ScanRoutesFolder Then
        For Each objFile In m_colRouteFiles
            Call chRoute.AddChoice(objFile.FileName)
        Next objFile
    Else
        PrintErrorMessage "UpdateRoutesList : failed to list route folder content"
        GoTo Fin
    End If
    
    'try to select the current route
    If g_Nav.Route.Description <> "" Then
        Dim i As Integer
        Dim bFound As Boolean
            
        For i = 0 To chRoute.ChoiceCount - 1
            If SameText(chRoute.Text(i), g_Nav.Route.Description) Then
                bFound = True
                chRoute.Selected = i
                Exit For
            End If
        Next i
        
        'if couldnt find the current route, create one
        If Not bFound Then
            Call SaveRouteAs(g_Nav.Route.Description)
        End If
    Else
        If chRoute.ChoiceCount > 0 Then
            chRoute.Selected = 0
        End If
    End If
    
Fin:
    Set objFile = Nothing
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsNav.UpdateRoutesList - " & Err.Description
    Resume Fin
End Sub


Private Function SaveRouteAs(ByVal sNewRouteName As String)
On Error GoTo ErrorHandler
    Dim bRet As Boolean
    
    m_sCurRoute = sNewRouteName
    g_Nav.Route.Description = m_sCurRoute
    If g_Nav.Route.SaveRoute(GetRouteFilePath(m_sCurRoute)) Then
        Call UpdateRoutesList
        PrintMessage "Route '" & m_sCurRoute & "' saved successfully."
    Else
        PrintErrorMessage "Failed to save the route."
    End If
    
Fin:
    SaveRouteAs = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsNav.SaveRouteAs(" & sNewRouteName & ")"
    Resume Fin
End Function


